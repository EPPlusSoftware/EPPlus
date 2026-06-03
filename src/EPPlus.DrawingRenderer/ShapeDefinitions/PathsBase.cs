namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public abstract class PathsBase
    {
        public abstract PathDrawingType Type { get; }

        internal abstract PathsBase Clone();
        public abstract double EndX { get; }
        public abstract double EndY { get; }
        public abstract void TranslateCoordiantesToPointsAndDegrees(double coordinateRatio, double angleRatio);
    }
}
