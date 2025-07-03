using System;
using OfficeOpenXml.PDF.Math;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfTransform
    {
        public static PdfTransform Identity => new PdfTransform(0, 0, 1, 1);

        public Vector2 Position { get; set; } = Vector2.Zero;

        public int Z { get; set; } = 0;

        public Vector2 Scale { get; set; } = Vector2.One;

        public Vector2 Size { get; set; } = Vector2.Zero;


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

        public PdfTransform(Vector2 position, Vector2 size)
            : this(position, size, Vector2.One, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 size, Vector2 scale)
            : this(position, size, scale, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 size, Vector2 scale, double rotation)
            : this(position, size, scale, rotation, null) { }

        public PdfTransform(Vector2 position, Vector2 size, PdfTransform parent)
            : this(position, size, Vector2.One, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 size, Vector2 scale, PdfTransform parent)
            : this(position, size, scale, 0d, null) { }

        public PdfTransform(Vector2 position, Vector2 size, Vector2 scale, double rotation, PdfTransform parent)
        {
            Position = position;
            Size = size;
            Scale = scale;
            Rotation = rotation;
            Parent = parent;
        }

        public PdfTransform(double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
        {
            Position = new Vector2(x, y);
            Size = new Vector2(width, height);
            Scale = new Vector2(scaleX, scaleY);
            Rotation = rotation;
            Parent = parent;
        }
    }
}
