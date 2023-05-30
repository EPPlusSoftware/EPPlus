namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal struct QuadRangeItem<T>
    {
        public QuadRangeItem(QuadRange range, T value)
        {
            Range=range;
            Value=value;
        }
        public QuadRange Range{ get; }
        public T Value { get; }
    }
}