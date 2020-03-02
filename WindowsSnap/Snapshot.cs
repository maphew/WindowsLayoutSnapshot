using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using static WindowsSnap.Native;

namespace WindowsSnap
{
    internal class Snapshot
    {
        private readonly string[] RestrictedWindows = {"explorer"};

        public Snapshot(bool userInitiated)
        {
            Logger.Console("New snapshot");
            EnumWindows(EvalWindow, 0);

            TimeTaken = DateTime.UtcNow;
            UserInitiated = userInitiated;

            var pixels = new List<long>();
            foreach (var screen in Screen.AllScreens)
            {
                pixels.Add(screen.Bounds.Width * screen.Bounds.Height);
            }
            MonitorPixelCounts = pixels.ToArray();
            NumMonitors = pixels.Count;

            Logger.Log($"Snapshot: {TimeTaken} ({(UserInitiated ? "user" : "auto")}) {NumMonitors} monitors");
        }

        public DateTime TimeTaken { get; }
        public bool UserInitiated { get; }
        public long[] MonitorPixelCounts { get; }
        public int NumMonitors { get; }
        public Dictionary<string, Window> HwndWindowDictionary { get; } = new Dictionary<string, Window>();
        public List<IntPtr> WindowsBackToTop { get; private set; } = new List<IntPtr>();

        internal TimeSpan Age => DateTime.UtcNow.Subtract(TimeTaken);

        /// <summary>
        ///     Create new snapshot.
        /// </summary>
        /// <param name="userInitiated">True if invoked by user, false if automated.</param>
        /// <returns>The new snapshot.</returns>
        internal static Snapshot TakeSnapshot(bool userInitiated)
        {
            return new Snapshot(userInitiated);
        }

        private bool EvalWindow(int hwndInt, int lParam)
        {
            var hwnd = new IntPtr(hwndInt);

            if (!IsAltTabWindow(hwnd))
            {
                return true;
            }

            // EnumWindows returns windows in Z order from back to front
            WindowsBackToTop.Add(hwnd);

            var winInfo = GetWindowInfoByHwnd(hwnd);
            var process = GetWindowProcessByHwnd(hwnd);
            var processId = process?.Id.ToString() ?? "?";
            var name = GetWindowName(process, hwnd);
            var text = GetWindowTextByHwnd(hwnd);
            var window = new Window(name, text, hwnd, process, winInfo[0], winInfo[1]);
            HwndWindowDictionary.Add(hwnd.ToString(), window);
      
            Logger.Console($" - {name} '{text}' ({hwnd}; {processId}; {process?.MainModule?.FileName ?? "?"}) {winInfo[0]}");

            return true;
        }

        private string GetWindowName(Process process, IntPtr hwnd)
        {
            if (process == null)
            {
                process = GetWindowProcessByHwnd(hwnd);
            }
            return process?.ProcessName ?? GetWindowTextByHwnd(hwnd);
        }

        private string GetWindowTextByHwnd(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return "";
            }

            string windowText = "";
            try
            {
                const int textLength = 256;
                var sb = new StringBuilder(textLength + 1);
                GetWindowText(hwnd, sb, sb.Capacity);
                windowText = sb.ToString();
            }
            catch (Exception ex)
            {
                Logger.Console($"Error: Could not find window text with hwnd {hwnd}: {ex}");
            }

            return windowText;
        }

        private Process GetWindowProcessByHwnd(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return null;
            }

            Window.GetWindowThreadProcessId(hwnd, out uint processId);
            try
            {
                return Process.GetProcessById((int)processId);
            }
            catch (Exception ex)
            {
                Logger.Console($"Error: Could not find process with hwnd {hwnd}: {ex}");
            }

            return null;
        }

        public string GetDisplayString()
        {
            var dt = TimeTaken.ToLocalTime();
            return dt.ToString("M") + ", " + dt.ToString("T");
        }

        internal void RestoreAndPreserveMenu(object sender, EventArgs e)
        {
            // ignore extra params
            // We save and restore the current foreground window because it's our tray menu
            // I couldn't find a way to get this handle straight from the tray menu's properties;
            //   the ContextMenuStrip.Handle isn't the right one, so I'm using win32
            // More info RE the restore is below, where we do it
            var currentForegroundWindow = GetForegroundWindow();

            try
            {
                Restore(sender, e);
            }
            finally
            {
                // A combination of SetForegroundWindow + SetWindowPos (via set_Visible) seems to be needed
                // This was determined by trying a bunch of stuff
                // This prevents the tray menu from closing, and makes sure it's still on top
                SetForegroundWindow(currentForegroundWindow);
                TrayIconForm.Cms.Visible = true;
            }
        }

        public void Restore()
        {
            Restore(this, EventArgs.Empty);
        }

        internal void Restore(object sender, EventArgs e)
        {
            try
            {
                // ignore extra params
                // first, restore the window rectangles and normal/maximized/minimized states
                Dictionary<string, Window> newWindows = new Dictionary<string, Window>();
                foreach (var hwndWindowPair in HwndWindowDictionary)
                {
                    IntPtr hwnd = new IntPtr(Convert.ToInt32(hwndWindowPair.Key, 16));
                    Window win = hwndWindowPair.Value;
                    if (RestrictedWindows.Contains(win.ProcessName))
                    {
                        Logger.Console($"Not moving restricted window: {win.Details}");
                        continue;
                    }
                    if (!TryMoveWindow(hwnd, win))
                    {
                        hwnd = FindHwndForDetachedWindow(win);
                        if (newWindows.ContainsKey(hwnd.ToString()))
                        {
                            Logger.Console($"Not moving window: Duplicate hwnd found for window {win.Details}: {hwnd}");
                            continue;
                        }

                        if (hwnd == IntPtr.Zero)
                        {
                            Logger.Console($"Could not find hwnd for window {win.Details}");
                            continue;
                        }

                        if (!TryMoveWindow(hwnd, win))
                        {
                            var lastErrorCode = GetLastError();
                            Logger.Log($"Can't move window {GetWindowName(null, hwnd)} ({hwnd}): {ResultWin32.GetErrorName(lastErrorCode)} ({lastErrorCode})");
                            continue;
                        }
                    }

                    if (!newWindows.ContainsKey(hwnd.ToString()))
                    {
                        newWindows.Add(hwnd.ToString(), win);
                    }
                    else
                    {
                        Logger.Console($"Duplicate hwnd found for window {win.Details}: {hwnd}");
                    }
                }

                // now update the z-orders
                WindowsBackToTop = WindowsBackToTop.FindAll(IsWindowVisible);
                var positionStructure = BeginDeferWindowPos(WindowsBackToTop.Count);
                for (var i = 0; i < WindowsBackToTop.Count; i++)
                {
                    positionStructure = UpdateZOrder(i, positionStructure);
                }

                HwndWindowDictionary.Clear();
                foreach (var kvp in newWindows)
                {
                    HwndWindowDictionary.Add(kvp.Key, kvp.Value);
                }

                EndDeferWindowPos(positionStructure);
            }
            catch (Exception ex)
            {
                Logger.Log($"Error restoring snapshot: {ex}");
            }
        }

        private IntPtr UpdateZOrder(int i, IntPtr positionStructure)
        {
            try
            {
                var hwnd = WindowsBackToTop[i];
                var hwndInsertAfter = i == 0 ? IntPtr.Zero : WindowsBackToTop[i - 1];
                var uFlags = DeferWindowPosCommands.SWP_NOMOVE
                             | DeferWindowPosCommands.SWP_NOSIZE
                             | DeferWindowPosCommands.SWP_NOACTIVATE;
                positionStructure = DeferWindowPos(positionStructure,
                                                   hwnd,
                                                   hwndInsertAfter,
                                                   0, 0, 0, 0,
                                                   uFlags);
            }
            catch (Exception ex)
            {
                Logger.Log($"Error updating z-order: {ex}");
            }

            return positionStructure;
        }

        private bool TryMoveWindow(IntPtr hwnd, Window win)
        {
            if (hwnd == IntPtr.Zero)
            {
                return false;
            }
            // make sure window will be inside a monitor
            var newpos = GetRectInsideNearestMonitor(win);
            return SetWindowPos(hwnd, 0, newpos.Left, newpos.Top, newpos.Width, newpos.Height,
                              0x0004 /*NOZORDER*/);
        }

        private Window GetWindowByHwnd(IntPtr hwnd)
        {
            foreach (string key in HwndWindowDictionary.Keys)
            {
                var intptr = new IntPtr(Convert.ToInt32(key, 16));
                if (intptr == hwnd && HwndWindowDictionary.TryGetValue(key, out Window win))
                {
                    return win;
                }
            }

            return null;
        }

        private static Rectangle GetRectInsideNearestMonitor(Window window)
        {
            Rectangle real = window.WindowPosition.CloneAsRectangle();
            Rectangle rect = window.WindowVisible.CloneAsRectangle();
            Rectangle monitorRect = Screen.GetWorkingArea(rect); // use workspace coordinates
            var y = new Rectangle(
                Math.Max(monitorRect.Left, Math.Min(monitorRect.Right - rect.Width, rect.Left)),
                Math.Max(monitorRect.Top, Math.Min(monitorRect.Bottom - rect.Height, rect.Top)),
                Math.Min(monitorRect.Width, rect.Width),
                Math.Min(monitorRect.Height, rect.Height)
            );
            if (rect != real) // support different real and visible position
                y = new Rectangle(
                    y.Left - rect.Left + real.Left,
                    y.Top - rect.Top + real.Top,
                    y.Width - rect.Width + real.Width,
                    y.Height - rect.Height + real.Height
                );

            if (y != real)
            {
                Logger.Console("Moving " + real + "→" + y + " in monitor " + monitorRect);
            }

            return y;
        }

        private static bool IsAltTabWindow(IntPtr hwnd)
        {
            if (!IsWindowVisible(hwnd))
            {
                return false;
            }

            var extendedStyles = GetWindowLongPtr(hwnd, -20); // GWL_EXSTYLE
            if ((extendedStyles.ToInt64() & WS_EX_APPWINDOW) > 0)
            {
                return true;
            }
            if ((extendedStyles.ToInt64() & WS_EX_TOOLWINDOW) > 0)
            {
                return false;
            }

            var hwndTry = GetAncestor(hwnd, GetAncestor_Flags.GetRootOwner);
            var hwndWalk = IntPtr.Zero;
            while (hwndTry != hwndWalk)
            {
                hwndWalk = hwndTry;
                hwndTry = GetLastActivePopup(hwndWalk);
                if (IsWindowVisible(hwndTry))
                {
                    break;
                }
            }

            return hwndWalk == hwnd;
        }

        private static WinRect[] GetWindowInfoByHwnd(IntPtr hwnd)
        {
            var winPos = new WinRect();
            var winVis = new WinRect();
            RECT pos;
            if (!GetWindowRect(hwnd, out pos))
            {
                throw new Exception("Error getting window rectangle");
            }
            winPos.CopyFrom(pos.ToRectangle());
            winVis.CopyFrom(pos.ToRectangle());

            var dwAttribute = 9; //DwmwaExtendedFrameBounds
            if (Environment.OSVersion.Version.Major >= 6
                && DwmGetWindowAttribute(hwnd, dwAttribute, out pos, Marshal.SizeOf(typeof(RECT))) == 0)
            {
                winVis.CopyFrom(pos.ToRectangle());
            }

            return new[] {winPos, winVis};
        }

        private IntPtr FindHwndForDetachedWindow(Window window)
        {
            try
            {
                List<Process> processes = Process.GetProcesses().ToList();
                if (string.IsNullOrWhiteSpace(window.ProcessPath))
                {
                    return IntPtr.Zero;
                }
                Process[] matchingProcesses = processes.Where(proc =>
                                                                  proc.ProcessName.Equals(window.ProcessName)
                                                                  && (proc.MainModule?.FileName?.Equals(window.ProcessPath) ?? false)).ToArray();
                if (!matchingProcesses.Any())
                {
                    Logger.Console($"No matching processes for window {window.Details}");
                    return IntPtr.Zero;
                }

                if (matchingProcesses.Length == 1)
                {
                    Logger.Console($"Found matching process for window {window.Details}");
                    return matchingProcesses.First().MainWindowHandle;
                }

                if (IsChrome(window.ProcessName))
                {
                    return HandleSpecialProcessWithManySameNamedChildren(window, matchingProcesses, IsChrome);
                }
                if (IsTeams(window.ProcessName))
                {
                    return HandleSpecialProcessWithManySameNamedChildren(window, matchingProcesses, IsTeams);
                }

                var matchText = new[]{window.Title};
                if (window.ProcessName.Equals("devenv")) // Include (Running) or not
                {
                    matchText.Append(window.Title.Contains(" (Running)")
                                         ? window.Title.Replace(" (Running)", "")
                                         : $"{window.Title} (Running)");
                }

                // Multiple processes have this name and path. Try to pick it out by the title.
                var matchesText = string.Join("\n   ", matchingProcesses.Select(p => p.ProcessName));
                Logger.Console($"Multiple matching processes for window {window.Details}\n   {matchesText})");
                var match = matchingProcesses.FirstOrDefault(proc => matchText.Contains(GetWindowTextByHwnd(proc.MainWindowHandle)))
                                             ?.MainWindowHandle ?? IntPtr.Zero;
                Logger.Console($"Multiple matching processes for window {window.Details}\n   selected {match}: {GetWindowTextByHwnd(match)})");
                return match;
            }
            catch (Exception ex)
            {
                Logger.Log($"Error finding detached window hwnd: {ex}");
                return IntPtr.Zero;
            }
        }

        private Process GetParentProcess(Process child)
        {
            try
            {
                var query = $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {child.Id}";
                var search = new ManagementObjectSearcher("root\\CIMV2", query);
                var results = search.Get().GetEnumerator();
                results.MoveNext();
                var queryObj = results.Current;
                var parentId = (uint) queryObj["ParentProcessId"];
                return Process.GetProcessById((int) parentId);
            }
            catch (Exception ex)
            {
                Logger.Console($"Error finding parent process: {ex}");
                return null;
            }
        }

        private IntPtr HandleSpecialProcessWithManySameNamedChildren(Window window,
                                                                     Process[] matchingProcesses,
                                                                     Func<string, bool> isSpecialProcess)
        {
            if (isSpecialProcess(window.ProcessName))
            {
                // chrome is a special case. Get the parent chrome.
                foreach (var chrome in matchingProcesses)
                {
                    if (!isSpecialProcess(GetParentProcess(chrome)?.ProcessName ?? "")) // parent isn't chrome, so is parent chrome
                    {
                        return chrome.MainWindowHandle;
                    }
                }
            }
            return IntPtr.Zero;
        }

        private bool IsChrome(string processName)
        {
            return processName?.Equals("chrome") ?? false;
        }

        private bool IsTeams(string processName)
        {
            return processName?.Equals("Teams") ?? false;
        }
    }
}