using EPPlus.Graphics.Geometry;
using System.Threading.Tasks;

namespace EPPlus.Graphics.Primitives
{
    internal class PrimitiveCircle
    {
        internal Vector2 Center;
        internal double Radius;

        internal PrimitiveCircle(Vector2 center, double radius)
        {
            Center = center;
            Radius = radius;
        }

        internal Vector2 GetPointOnCircle(double degrees)
        {
            return GeoMathUtils.GetSwingPointAtDegrees(degrees, Radius, Center);
        }
    }
}
