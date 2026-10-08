using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EPPlus.Graphics.Geometry;
using EPPlus.Graphics.Primitives;

namespace EPPlus.Graphics.TransformPrimitives
{
    //Position refers to the center of the circle
    public class TransformCircle : Transform
    {
        public PrimitiveCircle Circle;

        protected double Radius { get { return Circle.Radius; } }

        internal TransformCircle(double radius) : base()
        {
            Circle = new PrimitiveCircle(Vector2.Zero, radius);
        }

        internal TransformCircle(Transform parent, double radius) : base(Vector2.Zero, Vector2.One, parent)
        {
            Circle = new PrimitiveCircle(Vector2.Zero, radius);
        }

        internal Vector2 GetPointOnCircleGlobal(double degrees)
        {
            var ptOnCircle = Circle.GetPointOnCircle(degrees);
            return new Vector2(Position.X + ptOnCircle.X, Position.Y + ptOnCircle.Y);
        }

        internal Vector2 GetPointOnCircleLocal(double degrees)
        {
            var ptOnCircle = Circle.GetPointOnCircle(degrees);
            return new Vector2(LocalPosition.X + ptOnCircle.X, LocalPosition.Y + ptOnCircle.Y);
        }


        //TranformPoint CreateChildPointOnCircleAt(double degrees)
        //{
        //    var pointOnCircleGlobal = GetPointOnCircle(degrees);

        //    var point = new TranformPoint(pointOnCircleGlobal.X, pointOnCircleGlobal.Y);
        //    point.Parent = this;
        //    point.Name = Name + $"_At{degrees}";

        //    return point;
        //}
    }
}
