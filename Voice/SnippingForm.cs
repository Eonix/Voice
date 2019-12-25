using System;
using System.Drawing;
using System.Windows.Forms;
using Tesseract;

namespace Voice
{
    public partial class SnippingForm : Form
    {
        public Action<string> SnippingFinished;

        private Point startingPoint;
        private Rectangle selectionRectangle;

        public SnippingForm(Rectangle screenBounds)
        {
            InitializeComponent();
            SetBounds(screenBounds.X, screenBounds.Y, screenBounds.Width, screenBounds.Height);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
                HideForm();

            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            startingPoint = e.Location;
            selectionRectangle = new Rectangle(e.Location, Size.Empty);

            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            var smallestX = Math.Min(e.X, startingPoint.X);
            var smallestY = Math.Min(e.Y, startingPoint.Y);
            var largestX = Math.Max(e.X, startingPoint.X);
            var largestY = Math.Max(e.Y, startingPoint.Y);

            selectionRectangle = new Rectangle(smallestX, smallestY, largestX - smallestX, largestY - smallestY);

            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (selectionRectangle.Width < 1 || selectionRectangle.Height < 1)
            {
                HideForm();
                return;
            }

            var selectedImage = new Bitmap(selectionRectangle.Width, selectionRectangle.Height);
            var horizontalScale = BackgroundImage.Width / (double)Width;
            var verticalScale = BackgroundImage.Height / (double)Height;

            using (var graphics = Graphics.FromImage(selectedImage))
            {
                graphics.DrawImage(BackgroundImage,
                    new Rectangle(0, 0, selectedImage.Width, selectedImage.Height),
                    new Rectangle(
                        (int)(selectionRectangle.X * horizontalScale),
                        (int)(selectionRectangle.Y * verticalScale),
                        (int)(selectionRectangle.Width * horizontalScale),
                        (int)(selectionRectangle.Height * verticalScale)
                    ),
                    GraphicsUnit.Pixel
                );
            }

            SnippingFinished?.Invoke(ReadTextFromImage(selectedImage));

            HideForm();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (var brush = new SolidBrush(Color.FromArgb(120, Color.White)))
            {
                var sourceX = selectionRectangle.X;
                var targetX = selectionRectangle.X + selectionRectangle.Width;
                var sourceY = selectionRectangle.Y;
                var targetY = selectionRectangle.Y + selectionRectangle.Height;

                e.Graphics.FillRectangle(brush, new Rectangle(0, 0, sourceX, Height));
                e.Graphics.FillRectangle(brush, new Rectangle(targetX, 0, Width - targetX, Height));
                e.Graphics.FillRectangle(brush, new Rectangle(sourceX, 0, targetX - sourceX, sourceY));
                e.Graphics.FillRectangle(brush, new Rectangle(sourceX, targetY, targetX - sourceX, Height - targetY));
            }

            using (var pen = new Pen(Color.Red, 2))
            {
                e.Graphics.DrawRectangle(pen, selectionRectangle);
            }
        }

        private string ReadTextFromImage(Bitmap image)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                using (var pix = PixConverter.ToPix(image))
                using (var page = engine.Process(pix))
                {
                    return page.GetText();
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private void HideForm()
        {
            startingPoint = Point.Empty;
            selectionRectangle = new Rectangle();
            Hide();
        }
    }
}
