namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class ClosePath : PathsBase
    {
        public ClosePath()
        {

        }
        public override PathDrawingType Type => PathDrawingType.Close;
        internal override PathsBase Clone()
        {
            return new ClosePath();
        }
        public override double EndX => double.MinValue;
        public override double EndY => double.MinValue;
        public override void TranslateCoordiantesToPointsAndDegrees(double coordinateRatio, double angleRatio)
        {
        }
    }
}
