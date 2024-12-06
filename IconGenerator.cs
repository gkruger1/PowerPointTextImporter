using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PowerPointTextImporter
{
    public static class IconGenerator
    {
        public static Icon CreateIcon()
        {
            const int size = 32;
            using (var bitmap = new Bitmap(size, size))
            {
                using (var g = Graphics.FromImage(bitmap))
                {
                    // Set high quality rendering
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    // Draw orange background
                    using (var brush = new SolidBrush(Color.FromArgb(255, 209, 71, 0)))
                    {
                        g.FillRectangle(brush, 0, 0, size, size);
                    }

                    // Draw white "P"
                    using (var font = new Font("Arial", size * 0.6f, FontStyle.Bold))
                    using (var brush = new SolidBrush(Color.White))
                    {
                        var format = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString("P", font, brush, new RectangleF(0, 0, size, size), format);
                    }
                }

                // Convert bitmap to icon
                IntPtr hIcon = bitmap.GetHicon();
                try
                {
                    return Icon.FromHandle(hIcon);
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
