namespace EPPlus.DrawingRenderer
{
    /// <summary>
    /// Drawing type
    /// </summary>
    public enum PathDrawingType
    {
        /// <summary>
        /// Drawing-Path Move to command
        /// </summary>
        MoveTo,
        /// <summary>
        /// Drawing-Path Line to command
        /// </summary>
        LineTo,
        /// <summary>
        /// Drawing-Path Arc command
        /// </summary>
        ArcTo,
        /// <summary>
        /// Drawing-Path Cubic Berzier Curve command
        /// </summary>
        CubicBezierTo,
        /// <summary>
        /// Drawing-Path Quad Berier Curve command
        /// </summary>
        QuadBezierTo,
        /// <summary>
        /// Drawing-Path Close command
        /// </summary>
        Close
    }
}
