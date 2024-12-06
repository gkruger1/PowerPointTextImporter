using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace PowerPointTextImporter
{
    public static class IconResource
    {
        public static void GenerateIcon()
        {
            string resourcesDir = Path.Combine(Application.StartupPath, "Resources");
            Directory.CreateDirectory(resourcesDir);

            using (var bitmap = new Bitmap(32, 32))
            using (var g = Graphics.FromImage(bitmap))
            {
                // Set high quality rendering
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                // Set background color (PowerPoint orange)
                g.Clear(Color.FromArgb(255, 209, 71, 0));

                // Draw "P" in white
                using (var font = new Font("Arial", 24, FontStyle.Bold))
                using (var brush = Brushes.White)
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                })
                {
                    var rect = new RectangleF(0, 0, 32, 32);
                    g.DrawString("P", font, brush, rect, format);
                }

                // Save as icon
                string iconPath = Path.Combine(resourcesDir, "app.ico");
                using (var icon = Icon.FromHandle(bitmap.GetHicon()))
                using (var fs = File.Create(iconPath))
                {
                    icon.Save(fs);
                }
            }
        }
    }
}
