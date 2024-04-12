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

        internal const string CircleSvg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><circle style=' fill: {0}; stroke: {1}; stroke-width: 6%;' stroke-opacity='100%' cx='50%' cy='50%' r='43%'/></svg>";
        internal const string TrafficLightSvg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><circle style=' fill: {0}; stroke: #505050ff; stroke-width: 32%;' stroke-opacity='100%' cx='50%' cy='50%' r='50%'/></svg>";
        internal const string RedDiamond = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><rect x='40' y='-30' width='65%' height='65%' rx='1.9219' ry='1.9219' fill='#d65532' fill-rule='evenodd' stroke='#a95139' stroke-width='3%' transform='rotate(45)'/></svg>";
        internal const string YellowTriangle = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'> <path transform='translate(-16.5 6) scale(2.5 2.5)' d='m42.024 35.186-32.498-0.080195a2.0234 2.0234 60.141 0 1-1.7448-3.0394l16.704-28.768a1.4615 1.4615 0.14139 0 1 2.5313 0.00625l16.664 29.029a1.9044 1.9044 120.14 0 1-1.6563 2.8525z' fill='#eac282' fill-rule='evenodd' stroke='#a78433' stroke-width='1%'/></svg>";
        internal const string FlagBase = "<svg version='1.1' viewBox='0 0 100 100' xml:space='preserve' xmlns='http://www.w3.org/2000/svg'><g transform='translate(-13.143 -18.946) scale(1.5 1.5)'><g fill-rule='evenodd'><rect x='26.143' y='20.952' width='4.4309' height='56.526' fill='#808080'/><path transform='matrix(0 .76079 -.94646 0 92.688 12.868)' d='m31.903 28.105 20.802 35.591-41.224 0.21927z' fill='{0}' stroke='{1}'/></g></g></svg>";
        internal const string smallTriangleBase = "<svg version='1.1' viewBox='0 0 50 50' xmlns='http://www.w3.org/2000/svg'><path transform='rotate({0})' d='m25.203 8.7907c-11.807 0-23.614 0.073541-23.733 0.22066-0.23733 0.29424 23.258 29.424 23.733 29.424s23.97-29.129 23.733-29.424c-0.11866-0.14712-11.926-0.22066-23.733-0.22066z' fill='{1}' stroke='{2}' fill-rule='evenodd' stroke-width='2.1167'/></svg>";
        internal const string RedCross = "<svg width='100%' height='100%' version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><path  transform='translate(6 10)  scale(2 2)' d='m9.0937 1.0461-7.8717 7.6384 13.48 13.889-13.656 14.077 7.8659 7.6385 13.429-13.844 13.241 13.645 7.8715-7.6327-13.468-13.883 13.286-13.696-7.866-7.6384-13.064 13.463z' fill='#d76244' fill-rule='evenodd' stroke='#a23c20' stroke-width='1.4795'/></svg>";
        //0 id, 1 fillColor, 2 strokeColor, 3 radius

        internal static string GetIconSvg(eExcelconditionalFormattingCustomIcon icon)
        {
            string svg = GetIconSvgConvertedString(icon);
            return svg;
        }

        internal static string GetIconSvgConvertedString(eExcelconditionalFormattingCustomIcon icon)
        {
            string svg = GetIconSvgUnConvertedString(icon);
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(svg));
        }

        internal static string GetIconSvgUnConvertedString(eExcelconditionalFormattingCustomIcon icon)
        {
            //var wholeString = CircleFrontAndId
            //wholeString += GetIconMiddle(icon);
            //wholeString += string.Format(CircleSize, "50%", "50%", "50%");
            
            return GetIconMiddle(icon);
        }

        internal static string[] GetIconSetSvgs(string setString)
        {
            //var set = IconDict.GetIconSet(setString);

            //var retArr = new string[set.Length];
            var retArr = new string[2];

            retArr[1] = Convert.ToBase64String(Encoding.ASCII.GetBytes(GetIconMiddle(eExcelconditionalFormattingCustomIcon.GreenFlag)));
            retArr[0] = Convert.ToBase64String(Encoding.ASCII.GetBytes(RedCross));

            //for (int i = 0; i < set.Length; i++) 
            for (int i = 0; i < 1; i++)
            {
                // retArr[i] = GetIconSvg(set[i]);
            }

            return retArr;
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
                    return string.Format(CircleSvg, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowCircle:
                    return string.Format(CircleSvg, "#eac282", "#a88636");
                case eExcelconditionalFormattingCustomIcon.GreenCircle:
                    return string.Format(CircleSvg, "#68a490", "#387360");
                case eExcelconditionalFormattingCustomIcon.BlackCircle:
                case eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder:
                    return string.Format(CircleSvg, "#505050", "#33312f");
                case eExcelconditionalFormattingCustomIcon.GrayCircle:
                    return string.Format(CircleSvg, "#b1b1b1", "#74716e");
                case eExcelconditionalFormattingCustomIcon.PinkCircle:
                    return string.Format(CircleSvg, "#edb9ab", "#b18478");

                case eExcelconditionalFormattingCustomIcon.GreenTrafficLight:
                    return string.Format(TrafficLightSvg, "#68a490");
                case eExcelconditionalFormattingCustomIcon.YellowTrafficLight:
                    return string.Format(TrafficLightSvg, "#eac282");
                case eExcelconditionalFormattingCustomIcon.RedTrafficLight:
                    return string.Format(TrafficLightSvg, "#d86344");

                case eExcelconditionalFormattingCustomIcon.RedDiamond:
                    return RedDiamond;
                case eExcelconditionalFormattingCustomIcon.YellowTriangle:
                    return YellowTriangle;

                case eExcelconditionalFormattingCustomIcon.GreenFlag:
                    return string.Format(FlagBase, "#76a797", "#326f5b");
                case eExcelconditionalFormattingCustomIcon.YellowFlag:
                    return string.Format(FlagBase, "#eac282", "#a88636");
                case eExcelconditionalFormattingCustomIcon.RedFlag:
                    return string.Format(FlagBase, "#d65532", "#ac563e");

                case eExcelconditionalFormattingCustomIcon.RedDownTriangle:
                    return string.Format(smallTriangleBase, 0, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.GreenUpTriangle:
                    return string.Format(smallTriangleBase, 180, "#76a797", "#326f5b");

                case eExcelconditionalFormattingCustomIcon.RedCross:
                    return RedCross;

                default: 
                    throw new NotImplementedException();
            }
        }
    }
}
