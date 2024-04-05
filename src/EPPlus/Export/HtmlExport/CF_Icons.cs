using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class CF_Icons
    {
        internal const string CircleIcon = "<svg xmlns='http://www.w3.org/2000/svg'xmlns:svg='http://www.w3.org/2000/svg'> <circle id='{0}' style=' fill: {1}; stroke: {2}; stroke-width: 0.264583; stroke-opacity: 1;' cx='{3}' cy='{3}' r='{3}' /></svg>";
        internal const string DownArrow = "<?xml version='1.0' encoding='UTF-8' standalone='no'?> <svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg'> <g id='layer1' transform='rotate(180 7.925 9.5325)'> <path  id='arrow' style='fill: #d86344; stroke: #9e381c; stroke-width: 0.25;'  d='M 5.80786,0.125 V 11.57082 L 0.125,5.88848 v 5.19968 l 7.80004,7.80004 7.80005,-7.80004 V 5.88848 l -5.68286,5.68234 V 0.125 Z'/></g></svg>";
        //internal const string ArrowTransform = "transform='translate(-67.469 -28.575)' rotate({0})>";
        internal const string ArrowStart = "<?xml version='1.0' encoding='UTF-8' standalone='no'?> <svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg'> <g id='layer1' ";
        internal const string ArrowMiddle = " transform='rotate({0} 7.925 9.5325)'> <path  id='arrow' style='fill: {1}; stroke: {2}; stroke-width: 0.25;'";
        internal const string ArrowRotation = "transform='rotate({0} 7.925 9.5325)'";
        internal const string ArrowStyle = "style='fill: {0}; stroke: {1}; ";
        internal const string ArrowEnd = "stroke-width: 0.25;'  d='M 5.80786,0.125 V 11.57082 L 0.125,5.88848 v 5.19968 l 7.80004,7.80004 7.80005,-7.80004 V 5.88848 l -5.68286,5.68234 V 0.125 Z'/></g></svg>";

        internal const string CircleFrontAndId = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg'><circle id='{0}' ";
        internal const string CircleStyle = "style=' fill: {0}; stroke: {1}; stroke-width: 0.264583; stroke-opacity: 1;'";
        internal const string CircleSize = " cx='{0}' cy='{0}' r='{0}' /></svg>\r\n";

        //0 id, 1 fillColor, 2 strokeColor, 3 radius

        internal static string GetIconSvg(eExcelconditionalFormattingCustomIcon icon)
        {
            string svg = GetIconSvgConvertedString(icon);
            return $"background-image: url(data:image/svg+xml;base64,{svg})";
        }

        internal static string GetIconSvgConvertedString(eExcelconditionalFormattingCustomIcon icon)
        {
            string svg = GetIconSvgUnConvertedString(icon);
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(svg));
        }

        internal static string GetIconSvgUnConvertedString(eExcelconditionalFormattingCustomIcon icon)
        {
            var wholeString = string.Format(CircleFrontAndId, "circle1");
            wholeString += GetIconMiddle(icon);
            wholeString += string.Format(CircleSize, 12, 12, 12);
            return wholeString;
        }

        internal static string GetIconMiddle(eExcelconditionalFormattingCustomIcon icon)
        {
            switch (icon)
            {
                case eExcelconditionalFormattingCustomIcon.RedDownArrow:
                    return string.Format(ArrowMiddle, "0", "#d86344", "#9e381c");
                case eExcelconditionalFormattingCustomIcon.YellowSideArrow:
                    return string.Format(ArrowMiddle, "-90", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.GreenUpArrow:
                    return string.Format(ArrowMiddle, "180", "#76a797", "#3f7865");
                case eExcelconditionalFormattingCustomIcon.GrayDownArrow:
                    return string.Format(ArrowMiddle, "0", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GraySideArrow:
                    return string.Format(ArrowMiddle, "-90", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GrayUpArrow:
                    return string.Format(ArrowMiddle, "180", "#808080", "#646262");

                case eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow:
                    return string.Format(ArrowMiddle, "-45", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow:
                    return string.Format(ArrowMiddle, "-135", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow:
                    return string.Format(ArrowMiddle, "-45", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow:
                    return string.Format(ArrowMiddle, "-135", "#808080", "#646262");

                case eExcelconditionalFormattingCustomIcon.RedCircleWithBorder:
                case eExcelconditionalFormattingCustomIcon.RedCircle:
                    return string.Format(CircleStyle, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowCircle:
                    return string.Format(CircleStyle, "#eac282", "#a88636");
                case eExcelconditionalFormattingCustomIcon.GreenCircle:
                    return string.Format(CircleStyle, "#68a490", "#387360");
                case eExcelconditionalFormattingCustomIcon.BlackCircle:
                case eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder:
                    return string.Format(CircleStyle, "#505050", "#33312f");
                case eExcelconditionalFormattingCustomIcon.GrayCircle:
                    return string.Format(CircleStyle, "#b1b1b1", "#74716e");
                case eExcelconditionalFormattingCustomIcon.PinkCircle:
                    return string.Format(CircleStyle, "#edb9ab", "#b18478");

                default: throw new NotImplementedException();
            }
        }
    }
}
