using System;
using OfficeOpenXml.PDF.Math;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfTransform
    {
        public static PdfTransform Identity => new PdfTransform(0, 0, 1, 1);

        public Vector2 Position { get; set; } = Vector2.Zero;

        public Vector2 Scale { get; set; } = Vector2.One;

        private double rotationDegrees = 0;
        private double rotationRadians = 0;
        public double Rotation
        {
            get
            {
                return rotationDegrees;
            }
            set
            {
                rotationDegrees = value;
                rotationRadians = rotationDegrees * System.Math.PI / 100.0d;
            }
        }

        public PdfTransform Parent { get; set; }

        public PdfTransform() { }

        public PdfTransform(Vector2 position)
            : this(position, Vector2.One, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 scale)
            : this(position, scale, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 scale, double rotation)
            : this(position, scale, rotation, null) { }

        public PdfTransform(Vector2 position, PdfTransform parent)
            : this(position, Vector2.One, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 scale, PdfTransform parent)
            : this(position, scale, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 scale, double rotation, PdfTransform parent)
        {
            Position = position;
            Scale = scale;
            Rotation = rotation;
            Parent = parent;
        }

        public PdfTransform(double x, double y, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
        {
            Position = new Vector2(x, y);
            Scale = new Vector2(scaleX, scaleY);
            Rotation = rotation;
            Parent = parent;
        }
    }
}
