using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;

namespace OfficeOpenXml.PDF.Math
{
    /// <summary>
    /// Column Major Order 3x3 Matrix
    /// </summary>
    internal struct Matrix3x3
    {
        //[a d g]
        //[b e h]
        //[c f i]

        public double A = 0d;
        public double B = 0d;
        public double C = 0d;
        public double D = 0d;
        public double E = 0d;
        public double F = 0d;
        public double G = 0d;
        public double H = 0d;
        public double I = 0d;

        public static Matrix3x3 Identity => new Matrix3x3(1,0,0, 0,1,0, 0,0,1);

        /// <summary>
        /// Creates column-major order matrix. [c f i] = [0 0 1]
        /// [a d g]
        /// [b e h]
        /// [c f i]
        /// </summary>
        public Matrix3x3(double A, double B, double D, double E, double G, double H)
        {
            this.A = A;
            this.B = B;
            this.D = D;
            this.E = E;
            this.G = G;
            this.H = H;
            C = 0;
            F = 0;
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

        public Matrix3x3 Inverse()
        {
            //MatrixHelper uses Row-Major order, so we transpose the matrix.
            double[][] matrix = new double[3][] { new double[] { A, B, C }, new double[] { D, E, F }, new double[] { G, H, I } };
            var inverse = MatrixHelper.Inverse(matrix);
            //Return the transposed result so we get Column-Major matrix.
            return new Matrix3x3(inverse[0][0], inverse[1][0], inverse[2][0], inverse[0][1], inverse[1][1], inverse[2][1], inverse[0][2], inverse[1][2], inverse[2][2]);
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

        public static Matrix3x3 operator *(Matrix3x3 M1, Matrix3x3 M2)
        {
            //MatrixHelper uses Row-Major order, so we transpose the matrices.
            double[][] a = new double[3][] { new double[] { M1.A, M1.B, M1.C }, new double[] { M1.D, M1.E, M1.F }, new double[] { M1.G, M1.H, M1.I } };
            double[][] b = new double[3][] { new double[] { M2.A, M2.B, M2.C }, new double[] { M2.D, M2.E, M2.F }, new double[] { M2.G, M2.H, M2.I } };
            var result = MatrixHelper.Multiply(a, b);
            //Return the transposed result so we get Column-Major matrix.
            return new Matrix3x3(result[0][0], result[1][0], result[2][0], result[0][1], result[1][1], result[2][1], result[0][2], result[1][2], result[2][2]);
        }

        public static Vector2 operator *(Matrix3x3 M, Vector2 V)
        {
            return new Vector2
            (
                M.A * V.X + M.B * V.Y + M.C,
                M.D * V.X + M.E * V.Y + M.F
            );
        } 
    }
}
