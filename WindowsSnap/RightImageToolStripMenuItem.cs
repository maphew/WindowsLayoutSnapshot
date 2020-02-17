// Copyright (c) 2019 Cognex Corporation. All Rights Reserved

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using WindowsSnap.Properties;

namespace WindowsSnap
{
    /// <summary>
    ///     Custom tool strip menu item.
    /// </summary>
    public class RightImageToolStripMenuItem : ToolStripMenuItem
    {
        public RightImageToolStripMenuItem(string text)
            : base(text)
        {
        }

        public float[] MonitorSizes { get; set; }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var icon = Resources.monitor;
            var maxIconSizeScaling = (float) (e.ClipRectangle.Height - 8) / icon.Height;
            var maxIconSize = new Size((int) Math.Floor(icon.Width * maxIconSizeScaling),
                                       (int) Math.Floor(icon.Height * maxIconSizeScaling));
            var maxIconY = (int) Math.Round((e.ClipRectangle.Height - maxIconSize.Height) / 2f);

            var nextRight = e.ClipRectangle.Width - 5;
            for (var i = 0; i < MonitorSizes.Length; i++)
            {
                var thisIconSize = new Size((int) Math.Ceiling(maxIconSize.Width * MonitorSizes[i]),
                                            (int) Math.Ceiling(maxIconSize.Height * MonitorSizes[i]));
                var thisIconLocation = new Point(nextRight - thisIconSize.Width,
                                                 maxIconY + (maxIconSize.Height - thisIconSize.Height));

                // Draw with transparency
                var cm = new ColorMatrix {Matrix33 = 0.7f};
                // opacity
                using (var ia = new ImageAttributes())
                {
                    ia.SetColorMatrix(cm);

                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    e.Graphics.DrawImage(icon, new Rectangle(thisIconLocation, thisIconSize), 0, 0, icon.Width,
                                         icon.Height, GraphicsUnit.Pixel, ia);
                }

                nextRight -= thisIconSize.Width + 4;
            }
        }
    }
}