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
        /// BulletImage is at least as big as FontSize.
        /// </summary>
        AtLeast,
        /// <summary>
        /// BulletImage is no bigger than FontSize.
        /// </summary>
        AtMost
    }
}
