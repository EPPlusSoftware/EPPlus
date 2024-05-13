/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  15/04/2023         EPPlus Software AB         version 7.2
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class CF_Icons
    {
        internal const string SvgInitial = "<svg version='1.1' viewBox='0 0 {0} {1}' xmlns='http://www.w3.org/2000/svg'>";
        internal const string SvgIntialStandard = "<svg version='1.1' viewBox='0 0 100 100' xmlns='http://www.w3.org/2000/svg'>";

        internal const string ArrowNonViewbox = "<svg xmlns='http://www.w3.org/2000/svg'><g id='layer1' transform='rotate({0} 7.925 9.5325)'><path id='arrow' fill='{1}' stroke='{2}' stroke-width='0.25' d='M 5.80786,0.125 V 11.57082 L 0.125,5.88848 v 5.19968 l 7.80004,7.80004 7.80005,-7.80004 V 5.88848 l -5.68286,5.68234 V 0.125 Z'/></g></svg>";
        //Scaled down to ensure the rotated symbol fits as it's height is higher than its width.
        internal const string Arrow = "<path transform-origin='50% 50%' transform='rotate({0}) scale(0.8 0.8)' d='m30.529 0.6582-0.001953 60.18-29.797-29.797-0.072266 27.355 40.924 40.922 41.072-40.922v-27.355l-29.799 29.797v-60.18z' fill='{1}' stroke='{2}' stroke-width='3.2'/>";

        internal const string CircleSvg = "<circle style=' fill: {0}; stroke: {1}; stroke-width: 6%;' stroke-opacity='100%' cx='50%' cy='50%' r='43%'/>";
        internal const string TrafficLightSvg = "<circle style=' fill: {0}; stroke: #505050ff; stroke-width: 32%;' stroke-opacity='100%' cx='50%' cy='50%' r='50%'/>";
        internal const string RedDiamond = "<rect x='40' y='-30' width='65%' height='65%' rx='1.9219' ry='1.9219' fill='#d65532' stroke='#a95139' stroke-width='3%' transform='rotate(45)'/>";
        internal const string YellowTriangle = "<path transform='translate(-16.5 6) scale(2.5 2.5)' d='m42.024 35.186-32.498-0.080195a2.0234 2.0234 60.141 0 1-1.7448-3.0394l16.704-28.768a1.4615 1.4615 0.14139 0 1 2.5313 0.00625l16.664 29.029a1.9044 1.9044 120.14 0 1-1.6563 2.8525z' fill='#eac282' fill-rule='evenodd' stroke='#a78433' stroke-width='1%'/>";
        internal const string FlagBase = "<g transform='translate(-13.143 -18.946) scale(1.5 1.5)'><g><rect x='26.143' y='20.952' width='4.4309' height='56.526' fill='#808080'/><path transform='matrix(0 .76079 -.94646 0 92.688 12.868)' d='m31.903 28.105 20.802 35.591-41.224 0.21927z' fill='{0}' stroke='{1}'/></g></g>";
        internal const string SmallTriangleBase = "<path transform-origin='50% 50%' transform='rotate({0})' d='m49.984 31.455c-0.96622 0-48.796 59.296-48.312 59.895 0.48311 0.59895 96.14 0.59895 96.623 0s-47.344-59.895-48.311-59.895z' fill='{1}' stroke='{2}' stroke-width='3.2257'/>";

        internal const string RedCross = "<path d='m20.415 2.3307-17.689 16.848 30.291 30.635-30.687 31.049 17.676 16.848 30.176-30.535 29.754 30.096 17.689-16.836-30.266-30.623 29.857-30.209-17.676-16.848-29.358 29.694z' fill='{0}' stroke='{1}' stroke-width='3.2963' {2}/>";
        internal const string Exclamation = "<g fill='{0}' stroke='{1}' stroke-width='1.5863' {2}><ellipse cx='52.448' cy='89.226' rx='10.086' ry='9.738'/><path d='m40.836 0.79314h22.722l-3.9677 75.585h-14.792z'/></g>";
        internal const string Checkmark = "<path fill='{0}' stroke='{1}' stroke-width='3.2' {2} d='m80.752 1.9707-40.986 63.324-24.045-23.678-13.867 13.926 37.311 41.52c21.329-25.855 39.554-54.017 59.256-81.205z'/>";
        internal const string StarBase = "{0}<path d='m49.75 1.6191c-1.7923-1e-7 -11.105 34.102-12.555 35.207-1.45 1.1047-35.257-0.14682-35.811 1.6406-0.55384 1.7874 27.497 21.613 28.051 23.4 0.55384 1.7874-11.028 35.118-9.5781 36.223s28.1-20.744 29.893-20.744c1.7923 0 28.441 21.849 29.891 20.744 1.45-1.1047-10.132-34.435-9.5781-36.223 0.55384-1.7874 28.607-21.613 28.053-23.4-0.55384-1.7874-34.361-0.53592-35.811-1.6406-1.45-1.1047-10.762-35.207-12.555-35.207z' fill='{1}' stroke='#a6812d' stroke-width='3.1177'/>";
        internal const string YellowDash = "<rect x='1.7947' y='31.795' width='96.41' height='33.145' fill='#edc87e' stroke='#a4802c' stroke-width='3.5893'/>";
        
        internal const string SignalMeter = "<g stroke-width='3.2'>" +
            "<path {0} d='m0.60156 58.35v41.049h19.523v-41.049z'/>" +
            "<path {1} d='m21.207 38.609v60.664h19.273v-60.664z'/>" +
            "<path {2} d='m41.789 19.557v79.623h19.07v-79.623z'/>" +
            "<path {3} d='m62.357 0.91602v98.18h18.895v-98.18z'/></g>";
        
        internal const string FilledBoxes = "<g>" +
            "<rect id='middle-background' x='46.176' y='46.679' width='7.834' height='7.8854' rx='.73604' ry='.74507' fill='#757575' stroke-width='.35486'/>" +
            "<g stroke-width='2.2'>" +
            "<rect {0} x='.67963' y='50.233' width='48.757' height='49.077' rx='4.581' ry='4.6372'/>" +
            "<rect {1} x='50.562' y='50.234' width='48.757' height='49.077' rx='4.581' ry='4.6372'/>" +
            "<rect {2} x='.87471' y='.6901' width='48.757' height='49.077' rx='4.581' ry='4.6372'/>" +
            "<rect {3} x='50.563' y='.6892' width='48.757' height='49.077' rx='4.581' ry='4.6372'/></g></g>";

        internal const string AllQuarters = "<circle fill='#505050' cx='50' cy='50' r='43'/>" +
            "<g fill='#fff'>" +
            "<path {0}d='m6 49c0 23.748 21.252 45 45 45v-45z'/>" +
            "<path {1}d='m49 49v45c23.748 0 45-21.252 45-45z'/>" +
            "<path {2}d='m51 6c-23.748 0-45 21.252-45 45h45z'/>" +
            "<path {3}d='m49 6v45h45c0-23.748-21.252-45-45-45z'/>" +
            "<circle cx='50' cy='50' r='43%' fill-opacity='0' stroke='#33312f' stroke-width='6%'/></g>";

        internal const string Hide = "display='none' ";
        internal const string ShrinkAndCenter = "transform='translate(20 20) scale(0.6 0.6)'";

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
            string svg = GetIconMiddle(icon);
            if(Enum.GetName(typeof(eExcelconditionalFormattingCustomIcon), icon).Contains("Arrow"))
            {
                svg = AddParentNodeSvg(svg, 85);
            }
            else
            {
                svg = AddParentNodeSvg(svg);
            }
            return svg;
        }

        internal static string AddParentNodeSvg(string svg, int width= 100, int height = 100)
        {
            return string.Format(SvgInitial, width, height) + svg + "</svg>";
        }

        internal static string SetActiveIcons(int numActive, string SvgType)
        {
            var fill = "fill='{0}' stroke='{1}'";
            var inactiveFill = string.Format(fill, "#b3b3b3", "#757575");
            var activeFill = string.Format(fill, "#4d82b8", "#335f8c");

            var barFills = new string[4]
            {
                inactiveFill, inactiveFill, inactiveFill, inactiveFill
            };

            for(int i = 0; i < numActive; i++) 
            {
                barFills[i] = activeFill;
            }

            return string.Format(SvgType, barFills);
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

                case eExcelconditionalFormattingCustomIcon.RedFlag:
                    return string.Format(FlagBase, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowFlag:
                    return string.Format(FlagBase, "#eac282", "#a88636");
                case eExcelconditionalFormattingCustomIcon.GreenFlag:
                    return string.Format(FlagBase, "#76a797", "#326f5b");

                case eExcelconditionalFormattingCustomIcon.RedCircleWithBorder:
                case eExcelconditionalFormattingCustomIcon.RedCircle:
                    return string.Format(CircleSvg, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowCircle:
                    return string.Format(CircleSvg, "#eac282", "#a88636");
                case eExcelconditionalFormattingCustomIcon.GreenCircle:
                    return string.Format(CircleSvg, "#68a490", "#387360");

                case eExcelconditionalFormattingCustomIcon.RedTrafficLight:
                    return string.Format(TrafficLightSvg, "#d86344");
                case eExcelconditionalFormattingCustomIcon.YellowTrafficLight:
                    return string.Format(TrafficLightSvg, "#eac282");
                case eExcelconditionalFormattingCustomIcon.GreenTrafficLight:
                    return string.Format(TrafficLightSvg, "#68a490");

                case eExcelconditionalFormattingCustomIcon.RedDiamond:
                    return RedDiamond;
                case eExcelconditionalFormattingCustomIcon.YellowTriangle:
                    return YellowTriangle;


                case eExcelconditionalFormattingCustomIcon.RedCrossSymbol:
                    return GetIconMiddle(eExcelconditionalFormattingCustomIcon.RedCircle) +
                        string.Format(RedCross, "#fefdfdff", "#a23c20ff", ShrinkAndCenter);
                case eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol:
                    return GetIconMiddle(eExcelconditionalFormattingCustomIcon.YellowCircle) +
                        string.Format(Exclamation, "#fefdfdff", "a78433", ShrinkAndCenter);
                case eExcelconditionalFormattingCustomIcon.GreenCheckSymbol:
                    return GetIconMiddle(eExcelconditionalFormattingCustomIcon.GreenCircle) +
                        string.Format(Checkmark, "#fefdfdff", "#689485", ShrinkAndCenter);

                case eExcelconditionalFormattingCustomIcon.RedCross:
                    return string.Format(RedCross, "#d76244", "#a23c20", "");
                case eExcelconditionalFormattingCustomIcon.YellowExclamation:
                    return string.Format(Exclamation, "#edc87e", "#a78433", "");
                case eExcelconditionalFormattingCustomIcon.GreenCheck:
                    return string.Format(Checkmark, "#71a392", "#689485", "");

                case eExcelconditionalFormattingCustomIcon.HalfGoldStar:
                    return string.Format(StarBase, "<defs><linearGradient id='gradientFill'><stop offset='50%' stop-color='#eac282'/><stop offset='50%' stop-color='#ffffff'/></linearGradient></defs>", "url(#gradientFill)");
                case eExcelconditionalFormattingCustomIcon.SilverStar:
                    return string.Format(StarBase, "", "#ffffff");
                case eExcelconditionalFormattingCustomIcon.GoldStar:
                    return string.Format(StarBase, "", "#eac282");

                case eExcelconditionalFormattingCustomIcon.RedDownTriangle:
                    return string.Format(SmallTriangleBase, 180, "#d65532", "#ac563e");
                case eExcelconditionalFormattingCustomIcon.YellowDash:
                    return YellowDash;
                case eExcelconditionalFormattingCustomIcon.GreenUpTriangle:
                    return string.Format(SmallTriangleBase, 0, "#76a797", "#326f5b");

                case eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow:
                    return string.Format(Arrow, "-45", "#eac282", "#a4802b");
                case eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow:
                    return string.Format(Arrow, "-135", "#eac282", "#a4802b");

                case eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow:
                    return string.Format(Arrow, "-45", "#808080", "#646262");
                case eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow:
                    return string.Format(Arrow, "-135", "#808080", "#646262");

                case eExcelconditionalFormattingCustomIcon.BlackCircle:
                case eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder:
                    return string.Format(CircleSvg, "#505050", "#33312f");
                case eExcelconditionalFormattingCustomIcon.GrayCircle:
                    return string.Format(CircleSvg, "#b1b1b1", "#74716e");
                case eExcelconditionalFormattingCustomIcon.PinkCircle:
                    return string.Format(CircleSvg, "#edb9ab", "#b18478");

                case eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars:
                    return SetActiveIcons(0, SignalMeter);
                case eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar:
                    return SetActiveIcons(1,SignalMeter);
                case eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars:
                    return SetActiveIcons(2, SignalMeter);
                case eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars: 
                    return SetActiveIcons(3, SignalMeter);
                case eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars:
                    return SetActiveIcons(4, SignalMeter);

                case eExcelconditionalFormattingCustomIcon.WhiteCircle:
                    return string.Format(AllQuarters, "", "", "", "");
                case eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters:
                    return string.Format(AllQuarters, "", "", "", Hide);
                case eExcelconditionalFormattingCustomIcon.CircleWithTwoWhiteQuarters:
                    return string.Format(AllQuarters, "", Hide, "", Hide);
                case eExcelconditionalFormattingCustomIcon.CircleWithOneWhiteQuarter:
                    return string.Format(AllQuarters, Hide, Hide, "", Hide);

                case eExcelconditionalFormattingCustomIcon.ZeroFilledBoxes:
                    return SetActiveIcons(0, FilledBoxes);
                case eExcelconditionalFormattingCustomIcon.OneFilledBox:
                    return SetActiveIcons(1, FilledBoxes);
                case eExcelconditionalFormattingCustomIcon.TwoFilledBoxes:
                    return SetActiveIcons(2, FilledBoxes);
                case eExcelconditionalFormattingCustomIcon.ThreeFilledBoxes:
                    return SetActiveIcons(3, FilledBoxes);
                case eExcelconditionalFormattingCustomIcon.FourFilledBoxes:
                    return SetActiveIcons(4, FilledBoxes);

                case eExcelconditionalFormattingCustomIcon.NoIcon:
                    return "";

                default: 
                    throw new NotImplementedException($"the symbolId {(int)icon} with The symboltype: {Enum.GetName(typeof(eExcelconditionalFormattingCustomIcon), icon)} has not been implemented it is preceeded by {Enum.GetName(typeof(eExcelconditionalFormattingCustomIcon), icon-1)}");
            }
        }
    }
}
