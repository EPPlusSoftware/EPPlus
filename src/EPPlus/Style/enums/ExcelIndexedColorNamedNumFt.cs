namespace OfficeOpenXml.Style
{
    //See specify colors: https://support.microsoft.com/en-us/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5
    //The first 8 colors can be referred to by numberformats by name. e.g [Red]
    //The first 56 indexed colors by ColorX where X is the index. Including the named ones.
    public enum ExcelIndexedColorNamedNumFt
    {
        Black = ExcelIndexedColor.Indexed0,
        White = ExcelIndexedColor.Indexed1,
        Red = ExcelIndexedColor.Indexed2,
        Green = ExcelIndexedColor.Indexed3,
        Blue = ExcelIndexedColor.Indexed4,
        Yellow = ExcelIndexedColor.Indexed5,
        Magenta = ExcelIndexedColor.Indexed6,
        Cyan = ExcelIndexedColor.Indexed7
    }
}
