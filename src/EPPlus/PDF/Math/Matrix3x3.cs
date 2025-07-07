using OfficeOpenXml.RichData.IndexRelations;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.PDF.Math
{
    internal struct Matrix3x3
    {
        //[a b c]
        //[d e f]
        //[g h i]

        public double A;
        public double B;
        public double C = 0d;
        public double D;
        public double E;
        public double F = 0d;
        public double G;
        public double H;
        public double I = 1d;

        public static Matrix3x3 Identity => new Matrix3x3(1,0, 0,1, 0,0);

        public Matrix3x3(double A, double B, double D, double E, double G, double H)
        {
            this.A = A;
            this.B = B;
            this.D = D;
            this.E = E;
            this.G = G;
            this.H = H;
        }

        public static Matrix3x3 Translation(double tX, double tY) => new Matrix3x3(1, 0, 0, 1, tX, tY);

        public static Matrix3x3 Scaling(double sX, double sY) => new Matrix3x3(sX, 0, 0, sY, 0, 0);

        public static Matrix3x3 Rotation(double angleDegrees)
        {
            double radians = angleDegrees * System.Math.PI / 180d;
            double cos = System.Math.Cos(radians);
            double sin = System.Math.Sin(radians);
            return new Matrix3x3(cos, sin, -sin, cos, 0, 0);
        }

        public static Matrix3x3 operator *(Matrix3x3 M1, Matrix3x3 M2) //double check formula.
        {
            return new Matrix3x3(
                A: M1.A * M2.A + M1.D * M2.B,
                B: M1.B * M2.A + M1.E * M2.B,
                D: M1.A * M2.D + M1.D * M2.E,
                E: M1.B * M2.D + M1.E * M2.E,
                G: M1.A * M2.G + M1.D * M2.G + M1.H,
                H: M1.B * M2.G + M1.E * M2.H + M1.H
            );
        }

        // Transform a point
        public static Vector2 operator *(Matrix3x3 M, Vector2 V)
        {
            return new Vector2(
                M.A * V.X + M.D * V.Y + M.G,
                M.B * V.X + M.E * V.Y + M.H
            );
        }

    }
}
