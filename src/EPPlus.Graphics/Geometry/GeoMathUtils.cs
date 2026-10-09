using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Graphics.Geometry
{
    internal static class GeoMathUtils
    {
        internal static double DegreesToRadians(double degree)
        {
            return degree * (Math.Round((double)System.Math.PI, 14) / 180d);
        }

        internal static double RadiansToDegrees(double radians) => radians * (180d / (Math.Round((double)System.Math.PI, 14)));


        internal static Vector2 GetSwingPointAtDegrees(double degrees, double length, Vector2 rotationOriginPoint)
        {
            var angleRadians = DegreesToRadians(degrees);
            return GetSwingPointAtRadians(angleRadians, length, rotationOriginPoint);
        }

        /// <summary>
        /// rotate around the point with a length offset.
        /// useful for e.g. getting a point on a circle
        /// </summary>
        /// <param name="radians"></param>
        /// <param name="length"></param>
        /// <param name="rotationOriginPoint"></param>
        /// <returns></returns>
        internal static Vector2 GetSwingPointAtRadians(double radians, double length, Vector2 rotationOriginPoint)
        {
            var xPointGlobal = rotationOriginPoint.X + (length * Math.Cos(radians));
            var yPointGlobal = rotationOriginPoint.Y + (length * Math.Sin(radians));

            var swingPoint = new Vector2(xPointGlobal, yPointGlobal);

            return swingPoint;
        }
    }
}
