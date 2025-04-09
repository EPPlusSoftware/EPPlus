namespace OfficeOpenXml.Style
{
    //See specify colors: https://support.microsoft.com/en-us/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5
    //The first 8 colors can be referred to by numberformats by name. e.g [Red]
    //The first 56 indexed colors by ColorX where X is the index. Including the named ones.
    /// <summary>
    /// Get the color index by name.
    /// </summary>
    public enum ExcelIndexedColorNamedNumFt
    {
        /// <summary>
        /// Black
        /// </summary>
        Black = ExcelIndexedColor.Indexed0,
        /// <summary>
        /// White
        /// </summary>
        White = ExcelIndexedColor.Indexed1,
        /// <summary>
        /// Red
        /// </summary>
        Red = ExcelIndexedColor.Indexed2,
        /// <summary>
        /// Green
        /// </summary>
        Green = ExcelIndexedColor.Indexed3,
        /// <summary>
        /// Blue
        /// </summary>
        Blue = ExcelIndexedColor.Indexed4,
        /// <summary>
        /// Yellow
        /// </summary>
        Yellow = ExcelIndexedColor.Indexed5,
        /// <summary>
        /// Magenta
        /// </summary>
        Magenta = ExcelIndexedColor.Indexed6,
        /// <summary>
        /// Cyan
        /// </summary>
        Cyan = ExcelIndexedColor.Indexed7
    }
}
