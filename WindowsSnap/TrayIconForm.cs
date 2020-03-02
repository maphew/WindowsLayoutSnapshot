using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using static WindowsSnap.Native;
using Timer = System.Windows.Forms.Timer;

namespace WindowsSnap
{
    public partial class TrayIconForm : Form
    {

        private readonly IntPtr _notificationHandle;
        private readonly List<Snapshot> _snapshots = new List<Snapshot>();
        private readonly Timer _snapshotTimer = new Timer();
        private readonly TimeSpan _snapshotInterval = TimeSpan.FromMinutes(60);
        private readonly TimeSpan _snapshotMinInterval = TimeSpan.FromMinutes(5);
        
        private bool _firstStart = true;
        private Snapshot _menuShownSnapshot;
        private Padding? _originalTrayMenuArrowPadding;
        private Padding? _originalTrayMenuTextPadding;
        private DateTime _lastAutoSnapshot = DateTime.MinValue;

        public TrayIconForm()
        {
            try
            {
                InitializeComponent();
                Visible = false;

                _snapshotTimer.Interval = (int)_snapshotInterval.TotalMilliseconds;
                _snapshotTimer.Tick += SnapshotTimer_Tick;

                Cms = trayMenu;

                _notificationHandle = RegisterPowerSettingNotification(
                    Handle,
                    ref GUID_CONSOLE_DISPLAY_STATE,
                    DEVICE_NOTIFY_WINDOW_HANDLE
                );

                SystemEvents.SessionSwitch += SystemEventsOnSessionSwitch;

                _snapshots = Logger.ReadWindowsSnapJson();

                // Take first startup snapshot, if none read in
                if (_snapshots.Count < 1)
                {
                    TakeSnapshot(false);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error creating TrayIconForm: " + ex);
            }
        }

        internal static ContextMenuStrip Cms { get; set; }

        private void SystemEventsOnSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                _snapshotTimer.Enabled = false;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                _snapshotTimer.Enabled = true;
                _snapshots.Last().Restore();
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_POWERBROADCAST && m.WParam.ToInt32() == PBT_POWERSETTINGCHANGE)
            {
                var s = (POWERBROADCAST_SETTING) Marshal.PtrToStructure(m.LParam, typeof(POWERBROADCAST_SETTING));
                if (s.PowerSetting == GUID_CONSOLE_DISPLAY_STATE)
                    switch (s.Data)
                    {
                        case 0x00: // The display is off
                            _snapshotTimer.Enabled = false;
                            TakeSnapshot(false);
                            break;
                        case 0x01: // The display is on
                            _snapshotTimer.Enabled = true;

                            if (!_firstStart)
                            {
                                try
                                {
                                    // TODO: find better way, desktop isn't drawn yet when this event is received
                                    Thread.Sleep(5000);
                                    _snapshots.Last().Restore();
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log("Error restoring last snapshot: " + ex);
                                }
                            }

                            if (_firstStart)
                            {
                                _firstStart = false;
                            }

                            break;
                    }
            }

            base.WndProc(ref m);
        }

        private void SnapshotTimer_Tick(object sender, EventArgs e)
        {
            TakeSnapshot(false);
        }

        private void SnapshotToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Logger.Log("Menu item click initiated snapshot.");
            TakeSnapshot(true);
        }

        private void TakeSnapshot(bool userInitiated)
        {
            if (!userInitiated && DateTime.Now < _lastAutoSnapshot + _snapshotMinInterval)
            {
                return;
            }

            _snapshots.Add(Snapshot.TakeSnapshot(userInitiated));
            _lastAutoSnapshot = DateTime.Now;
            UpdateRestoreChoicesInMenu();
            try
            {
                Logger.WriteWindowsSnapJson(this._snapshots);
            }
            catch (Exception ex)
            {
                Logger.Log("Error writing JSON on snapshot: " + ex);
            }
        }

        private void ClearSnapshotsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _snapshots.Clear();
            UpdateRestoreChoicesInMenu();
        }

        private void JustNowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _menuShownSnapshot.Restore(null, EventArgs.Empty);
        }

        private void JustNowToolStripMenuItem_MouseEnter(object sender, EventArgs e)
        {
            SnapshotMousedOver(sender, e);
        }

        private void UpdateRestoreChoicesInMenu()
        {
            // construct the new list of menu items, then populate them
            // this function is idempotent

            var snapshotsOldestFirst = new List<Snapshot>(CondenseSnapshots(_snapshots, 8));
            var newMenuItems = new List<ToolStripItem>
            {
                quitToolStripMenuItem, aboutToolStripMenuItem, snapshotListEndLine
            };

            var maxNumMonitors = 0;
            var maxNumMonitorPixels = 0L;
            var showMonitorIcons = false;
            foreach (var snapshot in snapshotsOldestFirst)
            {
                if (maxNumMonitors != snapshot.NumMonitors && maxNumMonitors != 0)
                {
                    showMonitorIcons = true;
                }

                maxNumMonitors = Math.Max(maxNumMonitors, snapshot.NumMonitors);
                foreach (var monitorPixels in snapshot.MonitorPixelCounts)
                {
                    maxNumMonitorPixels = Math.Max(maxNumMonitorPixels, monitorPixels);
                }
            }

            foreach (var snapshot in snapshotsOldestFirst)
            {
                RightImageToolStripMenuItem menuItem = GetMenuUtemForSnapshot(snapshot, showMonitorIcons, maxNumMonitorPixels);
                newMenuItems.Add(menuItem);
            }

            newMenuItems.Add(justNowToolStripMenuItem);
            newMenuItems.Add(snapshotListStartLine);
            newMenuItems.Add(clearSnapshotsToolStripMenuItem);
            newMenuItems.Add(snapshotToolStripMenuItem);

            // If showing monitor icons, subtract 34 pixels from the right due to too much right padding.
            try
            {
                RemovePaddingForIcons(showMonitorIcons);
            }
            catch(Exception ex)
            {
                Logger.Log($"Error modifying padding: {ex}");
            }

            // If showing monitor icons, make the menu item width 50 + 22 * maxNumMonitors pixels wider to make room for the icons.
            try
            {
                AdjustPaddingForIcons(showMonitorIcons, maxNumMonitors);
            }
            catch(Exception ex)
            {
                Logger.Log($"Error modifying padding for icons: {ex}");
            }

            trayMenu.Items.Clear();
            trayMenu.Items.AddRange(newMenuItems.ToArray());
        }

        private void AdjustPaddingForIcons(bool showMonitorIcons, int maxNumMonitors)
        {
            var arrowPaddingField = typeof(ToolStripDropDownMenu).GetField("ArrowPadding",
                                                                           BindingFlags.NonPublic | BindingFlags.Static);
            if (arrowPaddingField == null)
            {
                Logger.Log("Error adjusting icon padding: arrow padding field is null.");
                return;
            }

            if (!_originalTrayMenuArrowPadding.HasValue)
            {
                _originalTrayMenuArrowPadding = (Padding) arrowPaddingField.GetValue(trayMenu);
            }

            var rightPadding = _originalTrayMenuArrowPadding.Value.Right + (showMonitorIcons ? 50 + 22 * maxNumMonitors : 0);
            arrowPaddingField.SetValue(trayMenu,
                                       new Padding(_originalTrayMenuArrowPadding.Value.Left,
                                                   _originalTrayMenuArrowPadding.Value.Top,
                                                   rightPadding,
                                                   _originalTrayMenuArrowPadding.Value.Bottom));
        }

        private void RemovePaddingForIcons(bool showMonitorIcons)
        {
            var textPaddingField =
                typeof(ToolStripDropDownMenu).GetField("TextPadding", BindingFlags.NonPublic | BindingFlags.Static);
            if (textPaddingField == null)
            {
                Logger.Log("Error removing right padding: text padding field is null.");
                return;
            }

            if (!_originalTrayMenuTextPadding.HasValue)
            {
                _originalTrayMenuTextPadding = (Padding) textPaddingField.GetValue(trayMenu);
            }

            var rightPadding = _originalTrayMenuTextPadding.Value.Right - (showMonitorIcons ? 34 : 0);
            textPaddingField.SetValue(trayMenu,
                                      new Padding(_originalTrayMenuTextPadding.Value.Left,
                                                  _originalTrayMenuTextPadding.Value.Top,
                                                  rightPadding,
                                                  _originalTrayMenuTextPadding.Value.Bottom));
        }

        private RightImageToolStripMenuItem GetMenuUtemForSnapshot(Snapshot snapshot, bool showMonitorIcons, long maxNumMonitorPixels)
        {
            var menuItem = new RightImageToolStripMenuItem(snapshot.GetDisplayString()) {Tag = snapshot};
            menuItem.Click += snapshot.Restore;
            menuItem.MouseEnter += SnapshotMousedOver;
            if (snapshot.UserInitiated) menuItem.Font = new Font(menuItem.Font, FontStyle.Bold);

            // monitor icons
            var monitorSizes = new List<float>();
            if (showMonitorIcons)
            {
                foreach (var monitorPixels in snapshot.MonitorPixelCounts)
                {
                    monitorSizes.Add((float) Math.Sqrt((float) monitorPixels / maxNumMonitorPixels));
                }
            }
            menuItem.MonitorSizes = monitorSizes.ToArray();
            return menuItem;
        }

        private List<Snapshot> CondenseSnapshots(List<Snapshot> snapshots, int maxNumSnapshots)
        {
            if (maxNumSnapshots < 2)
            {
                throw new Exception();
            }

            var y = new List<Snapshot>();
            y.AddRange(snapshots);
            y = RemoveOldSnapshots(maxNumSnapshots, y);
            y = RemoveRepetitiveSnapshots(maxNumSnapshots, y);
            return y;
        }

        private static List<Snapshot> RemoveRepetitiveSnapshots(int maxNumSnapshots, List<Snapshot> y)
        {
            // remove entries with the time most adjacent to another time
            while (y.Count > maxNumSnapshots)
            {
                var ixMostAdjacentNeighbors = -1;
                var lowestDistanceBetweenNeighbors = TimeSpan.MaxValue;
                for (var i = 1; i < y.Count - 1; i++)
                {
                    var distanceBetweenNeighbors = (y[i + 1].TimeTaken - y[i - 1].TimeTaken).Duration();

                    // a hack to make manual snapshots prioritized over automated snapshots
                    if (y[i].UserInitiated)
                    {
                        distanceBetweenNeighbors += TimeSpan.FromDays(1000000);
                    }

                    // a hack to make very recent snapshots prioritized over other snapshots
                    if (DateTime.UtcNow.Subtract(y[i].TimeTaken).Duration() <= TimeSpan.FromHours(2))
                    {
                        distanceBetweenNeighbors += TimeSpan.FromDays(2000000);
                    }

                    if (distanceBetweenNeighbors < lowestDistanceBetweenNeighbors)
                    {
                        lowestDistanceBetweenNeighbors = distanceBetweenNeighbors;
                        ixMostAdjacentNeighbors = i;
                    }
                }

                y.RemoveAt(ixMostAdjacentNeighbors);
            }

            return y;
        }

        private static List<Snapshot> RemoveOldSnapshots(int maxNumSnapshots, List<Snapshot> y)
        {
            // remove automatically-taken snapshots > 3 days old, or manual snapshots > 5 days old
            for (var i = 0; i < y.Count; i++)
            {
                if (y[i].Age > TimeSpan.FromDays(y[i].UserInitiated ? 5 : 3))
                {
                    y.RemoveAt(i);
                }

                if (y.Count <= maxNumSnapshots)
                {
                    break;
                }
            }

            return y;
        }

        private void SnapshotMousedOver(object sender, EventArgs e)
        {
            // We save and restore the current foreground window because it's our tray menu
            // I couldn't find a way to get this handle straight from the tray menu's properties;
            //   the ContextMenuStrip.Handle isn't the right one, so I'm using win32
            // More info RE the restore is below, where we do it
            var currentForegroundWindow = GetForegroundWindow();

            try
            {
                ((Snapshot) ((ToolStripMenuItem) sender).Tag).Restore(sender, e);
            }
            finally
            {
                // A combination of SetForegroundWindow + SetWindowPos (via set_Visible) seems to be needed
                // This was determined by trying a bunch of stuff
                // This prevents the tray menu from closing, and makes sure it's still on top
                SetForegroundWindow(currentForegroundWindow);
                trayMenu.Visible = true;
            }
        }

        private void QuitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnExit();
            Application.Exit();
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new About().Show();
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            _menuShownSnapshot = Snapshot.TakeSnapshot(false);
            justNowToolStripMenuItem.Tag = _menuShownSnapshot;

            // the context menu won't show by default on left clicks. We're going to have to ask it to show up.
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    // try using reflection to get to the private ShowContextMenu() function...which really 
                    // should be public but is not.
                    var showContextMenuMethod = trayIcon.GetType()
                                                        .GetMethod("ShowContextMenu",
                                                                   BindingFlags.NonPublic | BindingFlags.Instance);
                    if (showContextMenuMethod == null)
                    {
                        Logger.Log("Error handling left mouse click: ShowContextMenu is null.");
                        return;
                    }
                    showContextMenuMethod.Invoke(trayIcon, null);
                }
                catch (Exception ex)
                {
                    Logger.Log($"Error handling left mouse click: {ex}");
                }
            }
        }

        private void TrayIconForm_VisibleChanged(object sender, EventArgs e)
        {
            // Application.Run(Form) changes this form to be visible. Change it back.
            Visible = false;
        }

        private void OnExit()
        {
            try
            {
                Logger.WriteWindowsSnapJson(this._snapshots);
            }
            catch (Exception ex)
            {
                Logger.Log("Error with JSON on exit: " + ex);
            }
        }
    }
}