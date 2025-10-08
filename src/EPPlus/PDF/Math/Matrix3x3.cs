using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;

namespace OfficeOpenXml.PDF.Math
{
    /// <summary>
    /// Column Major Order 3x3 Matrix
    /// </summary>
    internal struct Matrix3x3
    {
        public double A = 0d;
        public double B = 0d;
        public double C = 0d;
        public double D = 0d;
        public double E = 0d;
        public double F = 0d;
        public double G = 0d;
        public double H = 0d;
        public double I = 0d;

        public static Matrix3x3 Identity => new Matrix3x3(1,0,0, 1,0,0, 0,0,1);

        /// <summary>
        /// Creates matrix
        /// [a b 0]
        /// [c d 0]
        /// [e f 1]
        /// </summary>
        public Matrix3x3(double A, double B, double C, double D, double E, double F)
        {
            this.A = A;
            this.B = B;
            this.C = C;
            this.D = D;
            this.E = E;
            this.F = F;
            G = 0;
            H = 0;
            I = 1;
        }

        public Matrix3x3(double A, double B, double C, double D, double E, double F, double G, double H, double I)
        {
            this.A = A;
            this.B = B;
            this.C = C;
            this.D = D;
            this.E = E;
            this.F = F;
            this.G = G;
            this.H = H;
            this.I = I;
        }

        public static Matrix3x3 Invert(Matrix3x3 m) => m.Inverse();

        public Matrix3x3 Inverse()
        {
            double[][] matrix = new double[3][]
            {
                new double[] { A, B, G },
                new double[] { C, D, H },
                new double[] { E, F, I }
            };
            var inverse = MatrixHelper.Inverse(matrix);
            return new Matrix3x3(
                inverse[0][0], inverse[0][1],
                inverse[1][0], inverse[1][1],
                inverse[2][0], inverse[2][1],
                inverse[0][2], inverse[1][2], inverse[2][2]
            );
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

        public Vector2 Transform(Vector2 p)
        {
            return new Vector2(
                A * p.X + B * p.Y + E,
                C * p.X + D * p.Y + F
            );
        }

        public static Matrix3x3 operator* (Matrix3x3 M1, Matrix3x3 M2)
        {
            double[][] a = new double[3][]
            {
                new double[] { M1.A, M1.B, M1.G },
                new double[] { M1.C, M1.D, M1.H },
                new double[] { M1.E, M1.F, M1.I }
            };
            double[][] b = new double[3][]
            {
                new double[] { M2.A, M2.B, M2.G },
                new double[] { M2.C, M2.D, M2.H },
                new double[] { M2.E, M2.F, M2.I }
            };
            var result = MatrixHelper.Multiply(a, b);
            return new Matrix3x3
            (
                result[0][0], result[0][1],
                result[1][0], result[1][1],
                result[2][0], result[2][1],
                result[0][2], result[1][2], result[2][2]
            );
        }

        public static Vector2 operator* ( Vector2 V, Matrix3x3 M)
        {
            return new Vector2
            (
                M.A * V.X + M.C * V.Y + M.E,
                M.B * V.X + M.D * V.Y + M.F
            );
        }
    }
}
