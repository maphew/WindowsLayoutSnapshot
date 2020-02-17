using System.Drawing;

namespace WindowsSnap
{
    public class WinRect
    {
        public int X { get; set; } = 0;
        public int Y { get; set; } = 0;
        public int Width { get; set; } = 0;
        public int Height { get; set; } = 0;

        public void CopyFrom(Rectangle rect)
        {
            X = rect.X;
            Y = rect.Y;
            Width = rect.Width;
            Height = rect.Height;
        }

        public Rectangle CloneAsRectangle()
        {
            return new Rectangle(X, Y, Width, Height);
        }

        public override string ToString()
        {
            return $"{X},{Y} {Width}x{Height}";
        }
    }
}
