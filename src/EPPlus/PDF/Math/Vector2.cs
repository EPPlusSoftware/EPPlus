namespace OfficeOpenXml.PDF.Math
{
    internal struct Vector2
    {
        public double X { get; set; } = 0;
        public double Y { get; set; } = 0;

        public Vector2() { }

        public Vector2(double x, double y)
        {
            X = x;
            Y = y;
        }

        public static readonly Vector2 Zero = new(0, 0);
        public static readonly Vector2 One = new(1, 1);

        public static Vector2 operator +(Vector2 v1, Vector2 v2) => new(v1.X + v2.X, v1.Y + v2.Y);
        public static Vector2 operator -(Vector2 v1, Vector2 v2) => new(v1.X - v2.X, v1.Y - v2.Y);
        public static Vector2 operator *(Vector2 v, double s) => new(v.X * s, v.Y * s);
        public static Vector2 operator *(double s, Vector2 v) => v * s;
    }
}
