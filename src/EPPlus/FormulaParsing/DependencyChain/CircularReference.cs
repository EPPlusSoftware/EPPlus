namespace OfficeOpenXml.FormulaParsing
{
    internal struct CircularReference
    {
        public CircularReference(ulong fromCell, ulong toCell)
        {
            FromCell = fromCell;
            ToCell = toCell;
        }
        internal ulong FromCell;
        internal ulong ToCell;
    }
}
