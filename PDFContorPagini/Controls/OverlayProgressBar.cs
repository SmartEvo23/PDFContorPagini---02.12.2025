using System;
using System.Drawing;
using System.Windows.Forms;

namespace PDFContorPagini.Controls
{
    public class OverlayProgressBar : ProgressBar
    {
        public OverlayProgressBar()
        {
            // enable custom painting and reduce flicker
            SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(BackColor);

            int width = ClientRectangle.Width;
            int height = ClientRectangle.Height;

            int range = (Maximum > Minimum) ? (Maximum - Minimum) : 1;
            double fraction = (Value - Minimum) / (double)range;
            int fillWidth = (int)Math.Round(fraction * width);

            // Draw empty area
            using (var backBrush = new SolidBrush(SystemColors.ControlLight))
                g.FillRectangle(backBrush, 0, 0, width, height);

            // Draw filled area (green)
            var filledRect = new Rectangle(0, 0, fillWidth, height);
            using (var fillBrush = new SolidBrush(Color.FromArgb(0, 150, 0)))
                g.FillRectangle(fillBrush, filledRect);

            // Prepare text
            string text = (Value >= Maximum) ? "Gata!" : $"{(int)Math.Round(fraction * 100)}%";

            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            using (var blackBrush = new SolidBrush(Color.Black))
            using (var whiteBrush = new SolidBrush(Color.White))
            using (var textFont = new Font(Font.FontFamily, 12f, Font.Style))
            {
                // Draw black text across whole control
                g.DrawString(text, textFont, blackBrush, ClientRectangle, sf);

                // Draw white text clipped to filled region so white appears only over green
                var state = g.Save();
                try
                {
                    g.SetClip(filledRect);
                    g.DrawString(text, textFont, whiteBrush, ClientRectangle, sf);
                }
                finally
                {
                    g.Restore(state);
                }
            }
        }
    }
}