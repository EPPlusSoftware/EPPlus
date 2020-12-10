namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Aspect ratio handling for a picture in a fill
    /// </summary>
    public enum eVmlAspectRatio
    {
        /// <summary>
        /// Ignore aspect issues. Default.
        /// </summary>
        Ignore,
        /// <summary>
        /// Image is at least as big as Size.
        /// </summary>
        AtLeast,
        /// <summary>
        /// Image is no bigger than Size.
        /// </summary>
        AtMost
    }
}
