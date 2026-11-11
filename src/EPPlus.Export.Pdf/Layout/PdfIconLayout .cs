using EPPlus.Graphics;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfIconLayout : Transform
    {
        public string IconName;

        public PdfIconLayout(double x, double y, double width, double height,
            double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y - height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 1;
        }
    }
}