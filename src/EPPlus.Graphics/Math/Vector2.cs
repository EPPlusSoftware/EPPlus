/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Graphics.Math
{
    /// <summary>
    /// This should REMAIN a struct
    /// Several math operations and the structure of parent-child hierarchy
    /// RELIES on that assumption. 
    /// As a class Zero and One can have x and y assigned to and this creates chaos as they should be constant
    /// </summary>
    internal struct Vector2
    {
        public double X { get; set; } = 0;
        public double Y { get; set; } = 0;

        public const double Z = 1; //Z is always 1 in pdf.

        public double Length
        {
            get { return System.Math.Sqrt(X * X + Y * Y); }
        }

        public double LengthSquared
        {
            get { return X * X + Y * Y; }
        }

        public Vector2() { }

        public Vector2(double x, double y)
        {
            X = x;
            Y = y;
        }

        public static readonly Vector2 Zero = new Vector2(0, 0);
        public static readonly Vector2 One = new Vector2(1, 1);

        public static Vector2 operator +(Vector2 v1, Vector2 v2) => new Vector2(v1.X + v2.X, v1.Y + v2.Y);
        public static Vector2 operator -(Vector2 v1, Vector2 v2) => new Vector2(v1.X - v2.X, v1.Y - v2.Y);
        public static Vector2 operator *(Vector2 v, double s)    => new Vector2(v.X * s, v.Y * s);
        public static Vector2 operator *(double s, Vector2 v)    => v * s;
        public static Vector2 operator *(Vector2 v1, Vector2 v2) => new Vector2(v1.X * v2.X, v1.Y * v2.Y);
        public static Vector2 operator /(Vector2 v1, Vector2 v2) => new Vector2(v1.X / v2.X, v1.Y / v2.Y);
        public static Vector2 operator /(Vector2 v, double s)    => new Vector2(v.X / s, v.Y / s);
        public static double Dot(Vector2 v1, Vector2 v2)
        {
            return v1.X * v2.X + v1.Y * v2.Y;
        }
        public static Vector2 Project(Vector2 vector, Vector2 onto)
        {
            double length = onto.LengthSquared;
            if (length == 0) return new Vector2(0, 0);
            double scale = Dot(vector, onto) / length;
            return new Vector2(onto.X * scale, onto.Y * scale);
        }
    }
}
