using System.Drawing;

namespace OfficeOpenXml.Utils
{
    internal static class ColorExtentions
    {
        internal static string To6CharHexString(this Color color)
        {
            return (color.ToArgb() & 0xFFFFFF).ToString("x6");
        }
    }

}
