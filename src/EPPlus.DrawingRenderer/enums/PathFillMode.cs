namespace EPPlus.DrawingRenderer
{
    /// <summary>
    /// How a shape path is filled.
    /// </summary>
    public enum PathFillMode
    {
        /// <summary>
        /// The corresponding path should have a normally shaded color applied to it’s fill
        /// </summary>
        Norm,
        /// <summary>
        /// The corresponding path should have a darker shaded color applied to it’s fill.
        /// </summary>
        Darken,
        /// <summary>
        /// The corresponding path should have a slightly darker shaded color applied to it’s fill.
        /// </summary>
        DarkenLess,
        /// <summary>
        /// The corresponding path should have a lightly shaded color applied to it’s fill.
        /// </summary>
        Lighten,
        /// <summary>
        /// The corresponding path should have a slightly lighter shaded color applied to it’s fill.
        /// </summary>
        LightenLess,
        /// <summary>
        /// The corresponding path should have no fill.
        /// </summary>
        None
    }
}
