using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class CF_Icons
    {
        internal const string ArrowNonViewbox = "<svg xmlns='http://www.w3.org/2000/svg'><g id='layer1' transform='rotate({0} 7.925 9.5325)'><path id='arrow' fill='{1}' stroke='{2}' stroke-width='0.25' d='M 5.80786,0.125 V 11.57082 L 0.125,5.88848 v 5.19968 l 7.80004,7.80004 7.80005,-7.80004 V 5.88848 l -5.68286,5.68234 V 0.125 Z'/></g></svg>";
        //Scaled down to ensure the rotated symbol fits as it's height is higher than its width.
        internal const string Arrow = "<svg version='1.1' viewBox='0 0 85 100' xmlns='http://www.w3.org/2000/svg'><path transform-origin='50% 50%' transform='rotate({0}) scale(0.8 0.8)' d='m30.529 0.6582-0.001953 60.18-29.797-29.797-0.072266 27.355 40.924 40.922 41.072-40.922v-27.355l-29.799 29.797v-60.18z' fill='{1}' stroke='{2}' stroke-width='3.2'/></svg>";

        internal const string CircleSvg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><circle style=' fill: {0}; stroke: {1}; stroke-width: 6%;' stroke-opacity='100%' cx='50%' cy='50%' r='43%'/></svg>";
        internal const string TrafficLightSvg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:svg='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><circle style=' fill: {0}; stroke: #505050ff; stroke-width: 32%;' stroke-opacity='100%' cx='50%' cy='50%' r='50%'/></svg>";
        internal const string RedDiamond = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><rect x='40' y='-30' width='65%' height='65%' rx='1.9219' ry='1.9219' fill='#d65532' fill-rule='evenodd' stroke='#a95139' stroke-width='3%' transform='rotate(45)'/></svg>";
        internal const string YellowTriangle = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'> <path transform='translate(-16.5 6) scale(2.5 2.5)' d='m42.024 35.186-32.498-0.080195a2.0234 2.0234 60.141 0 1-1.7448-3.0394l16.704-28.768a1.4615 1.4615 0.14139 0 1 2.5313 0.00625l16.664 29.029a1.9044 1.9044 120.14 0 1-1.6563 2.8525z' fill='#eac282' fill-rule='evenodd' stroke='#a78433' stroke-width='1%'/></svg>";
        internal const string FlagBase = "<svg version='1.1' viewBox='0 0 100 100' xml:space='preserve' xmlns='http://www.w3.org/2000/svg'><g transform='translate(-13.143 -18.946) scale(1.5 1.5)'><g fill-rule='evenodd'><rect x='26.143' y='20.952' width='4.4309' height='56.526' fill='#808080'/><path transform='matrix(0 .76079 -.94646 0 92.688 12.868)' d='m31.903 28.105 20.802 35.591-41.224 0.21927z' fill='{0}' stroke='{1}'/></g></g></svg>";
        internal const string SmallTriangleBase = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><path transform-origin='50% 50%' transform='rotate({0})' d='m49.984 31.455c-0.96622 0-48.796 59.296-48.312 59.895 0.48311 0.59895 96.14 0.59895 96.623 0s-47.344-59.895-48.311-59.895z' fill='{1}' stroke='{2}' stroke-width='3.2257'/></svg>";

        internal const string RedCross = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><path  transform='translate(6 10) scale(2 2)' d='m9.0937 1.0461-7.8717 7.6384 13.48 13.889-13.656 14.077 7.8659 7.6385 13.429-13.844 13.241 13.645 7.8715-7.6327-13.468-13.883 13.286-13.696-7.866-7.6384-13.064 13.463z' fill='#d76244' fill-rule='evenodd' stroke='#a23c20' stroke-width='1.4795'/></svg>";
        internal const string Exclamation = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><g transform='translate(24 0) scale(2,2)' fill='#edc87e' stroke='#a78433' stroke-width='.32'><path d='m5.6446 36.044c-2.5258 1.4e-5 -4.5734 1.977-4.5734 4.4158 2.598e-4 2.4385 2.0477 4.4152 4.5734 4.4152 2.5256-1.5e-5 4.5731-1.9767 4.5734-4.4152 3.6e-5 -2.4387-2.0475-4.4157-4.5734-4.4158z'/><path d='m0.37879 0.35967 1.7968 34.274h6.7076l1.7989-34.274z'/></g></svg>";
        internal const string Checkmark = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><path d='m80.752 1.9707-40.986 63.324-24.045-23.678-13.867 13.926 37.311 41.52c21.329-25.855 39.554-54.017 59.256-81.205z' fill='#71a392' stroke='#689485' stroke-width='3.2'/></svg>";
        internal const string StarBase = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'>{0}<path d='m49.75 1.6191c-1.7923-1e-7 -11.105 34.102-12.555 35.207-1.45 1.1047-35.257-0.14682-35.811 1.6406-0.55384 1.7874 27.497 21.613 28.051 23.4 0.55384 1.7874-11.028 35.118-9.5781 36.223s28.1-20.744 29.893-20.744c1.7923 0 28.441 21.849 29.891 20.744 1.45-1.1047-10.132-34.435-9.5781-36.223 0.55384-1.7874 28.607-21.613 28.053-23.4-0.55384-1.7874-34.361-0.53592-35.811-1.6406-1.45-1.1047-10.762-35.207-12.555-35.207z' fill='{1}' stroke='#a6812d' stroke-width='3.1177'/></svg>";
        internal const string YellowDash = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'><rect x='1.7947' y='31.795' width='96.41' height='33.145' fill='#edc87e' fill-rule='evenodd' stroke='#a4802c' stroke-width='3.5893'/></svg>";

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
            var set = IconDict.GetIconSet(setString);

            //var retArr = new string[set.Length];
            var retArr = new string[3];

            //retArr[0] = Convert.ToBase64String(Encoding.ASCII.GetBytes(StarBase));  
            //retArr[1] = Convert.ToBase64String(Encoding.ASCII.GetBytes(GetIconMiddle(eExcelconditionalFormattingCustomIcon.GreenFlag)));
            //retArr[2] = Convert.ToBase64String(Encoding.ASCII.GetBytes(GetIconMiddle(eExcelconditionalFormattingCustomIcon.RedCircle)));

            for (int i = 0; i < set.Length; i++) 
            {
                retArr[i] = GetIconSvg(set[i]);
            }

            //if(setString != "3Triangles")
            //{
            //    retArr[0] = Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format(SmallTriangleBaseFirst, 0, "#d65532", "#ac563e")));
            //    retArr[1] = GetIconSvg(eExcelconditionalFormattingCustomIcon.YellowDash);
            //    retArr[2] = Convert.ToBase64String(Encoding.ASCII.GetBytes(string.Format(SmallTriangleBaseFirst, 180, "#76a797", "#326f5b")));
            //}

            return retArr;
        }

        internal static string GetIconMiddle(eExcelconditionalFormattingCustomIcon icon)
        {
            switch (icon)
            {
                case eExcelconditionalFormattingCustomIcon.RedDownArrow:
                    return string.Format(Arrow, "0", "#d86344", "#9e381c");
                case eExcelconditionalFormattingCustomIcon.YellowSideArrow:
                    return string.Format(Arrow, "-90", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.GreenUpArrow:
                    return string.Format(Arrow, "180", "#76a797", "#3f7865");
                case eExcelconditionalFormattingCustomIcon.GrayDownArrow:
                    return string.Format(Arrow, "0", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GraySideArrow:
                    return string.Format(Arrow, "-90", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GrayUpArrow:
                    return string.Format(Arrow, "180", "#808080", "#646262");

                case eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow:
                    return string.Format(Arrow, "-45", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow:
                    return string.Format(Arrow, "-135", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow:
                    return string.Format(Arrow, "-45", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow:
                    return string.Format(Arrow, "-135", "#808080", "#646262");

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
                    return string.Format(SmallTriangleBase, 180, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowDash:
                    return YellowDash;
                case eExcelconditionalFormattingCustomIcon.GreenUpTriangle:
                    return string.Format(SmallTriangleBase, 0, "#76a797", "#326f5b");

                case eExcelconditionalFormattingCustomIcon.RedCross:
                    return RedCross;
                case eExcelconditionalFormattingCustomIcon.YellowExclamation:
                    return Exclamation;
                case eExcelconditionalFormattingCustomIcon.GreenCheck:
                    return Checkmark;

                case eExcelconditionalFormattingCustomIcon.GoldStar:
                    return string.Format(StarBase, "", "#eac282");
                case eExcelconditionalFormattingCustomIcon.HalfGoldStar:
                    return string.Format(StarBase, "<defs><linearGradient id='gradientFill'><stop offset='50%' stop-color='#eac282'/><stop offset='50%' stop-color='#ffffff'/></linearGradient></defs>", "url(#gradientFill)");
                case eExcelconditionalFormattingCustomIcon.SilverStar:
                    return string.Format(StarBase, "", "#ffffff");
                default: 
                    throw new NotImplementedException($"The symboltype: {Enum.GetName(typeof(eExcelconditionalFormattingCustomIcon), icon)} has not been implemented");
            }
        }
    }
}
