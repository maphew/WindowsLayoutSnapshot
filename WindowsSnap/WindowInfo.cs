using System.Drawing;

namespace WindowsSnap
{
    public class WindowInfo
    {
        public WinRect Position { get; set; } // real window border, we use this to move it
        public WinRect Visible { get; set; } // visible window borders, we use this to force inside a screen
    }
}
