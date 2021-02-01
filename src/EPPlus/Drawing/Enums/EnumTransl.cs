/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    internal static class EnumTransl
    {
        #region "Translate Enum functions"
        internal static string FromLineStyle(eLineStyle value)
        {
            string text = value.ToString();
            switch (value)
            {
                case eLineStyle.Dash:
                case eLineStyle.Dot:
                case eLineStyle.DashDot:
                case eLineStyle.Solid:
                    return text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1); //First to Lower case.
                case eLineStyle.LongDash:
                case eLineStyle.LongDashDot:
                case eLineStyle.LongDashDotDot:
                    return "lg" + text.Substring(4, text.Length - 4);
                case eLineStyle.SystemDash:
                case eLineStyle.SystemDashDot:
                case eLineStyle.SystemDashDotDot:
                case eLineStyle.SystemDot:
                    return "sys" + text.Substring(6, text.Length - 6);
                default:
                    throw (new Exception("Invalid Linestyle"));
            }
        }
        internal static eLineStyle ToLineStyle(string text)
        {
            switch (text)
            {
                case "dash":
                case "dot":
                case "dashDot":
                case "solid":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), text, true);
                case "lgDash":
                case "lgDashDot":
                case "lgDashDotDot":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), "Long" + text.Substring(2, text.Length - 2));
                case "sysDash":
                case "sysDashDot":
                case "sysDashDotDot":
                case "sysDot":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), "System" + text.Substring(3, text.Length - 3));
                default:
                    throw (new Exception("Invalid Linestyle"));
            }
        }
        internal static string FromLineCap(eLineCap value)
        {
            switch (value)
            {
                case eLineCap.Round:
                    return "rnd";
                case eLineCap.Square:
                    return "sq";
                default:
                    return "flat";
            }
        }
        internal static eLineCap ToLineCap(string text)
        {
            switch (text)
            {
                case "rnd":
                    return eLineCap.Round;
                case "sq":
                    return eLineCap.Square;
                default:
                    return eLineCap.Flat;
            }
        }
        internal static eCompundLineStyle     ToLineCompound(string s)
        {
            switch (s)
            {
                case "dbl":
                    return eCompundLineStyle.Double;
                case "sng":
                    return eCompundLineStyle.Single;
                case "thickThin":
                    return eCompundLineStyle.DoubleThickThin;
                case "thinThick":
                    return eCompundLineStyle.DoubleThinThick;
                default:
                    return eCompundLineStyle.TripleThinThickThin;
            }
        }

        internal static string FromLineCompound(eCompundLineStyle v)
        {
            switch (v)
            {
                case eCompundLineStyle.Double:
                    return "dbl";
                case eCompundLineStyle.Single:
                    return "sng";
                case eCompundLineStyle.DoubleThickThin:
                    return "thickThin";
                case eCompundLineStyle.DoubleThinThick:
                    return "thinThick";
                default:
                    return "tri";
            }
        }
        internal static ePenAlignment ToPenAlignment(string s)
        {
            if(s=="ctr")
            {
                return ePenAlignment.Center;
            }
            else
            {
                return ePenAlignment.Inset;
            }
        }

        internal static string FromPenAlignment(ePenAlignment v)
        {
            if(v==ePenAlignment.Center)
            {
                return "ctr";
            }
            else
            {
                return "in";
            }
        }

        #endregion

        internal static eUnderLineType TranslateUnderline(this string text)
        {
            switch (text)
            {
                case "sng":
                    return eUnderLineType.Single;
                case "dbl":
                    return eUnderLineType.Double;
                case "":
                    return eUnderLineType.None;
                default:
                    return text.ToEnum(eUnderLineType.None);
            }
        }
        internal static string TranslateUnderlineText(this eUnderLineType value)
        {
            switch (value)
            {
                case eUnderLineType.Single:
                    return "sng";
                case eUnderLineType.Double:
                    return "dbl";
                default:
                    string ret = value.ToString();
                    return ret.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + ret.Substring(1, ret.Length - 1);
            }
        }
        internal static eTextAnchoringType TranslateTextAchoring(this string text)
        {
            switch (text)
            {
                case "b":
                    return eTextAnchoringType.Bottom;
                case "ctr":
                    return eTextAnchoringType.Center;
                case "dist":
                    return eTextAnchoringType.Distributed;
                case "just":
                    return eTextAnchoringType.Justify;
                default:
                    return eTextAnchoringType.Top;
            }
        }
        internal static string TranslateTextAchoringText(this eTextAnchoringType value)
        {
            switch (value)
            {
                case eTextAnchoringType.Bottom:
                    return "b";
                case eTextAnchoringType.Center:
                    return "ctr";
                case eTextAnchoringType.Distributed:
                    return "dist";
                case eTextAnchoringType.Justify:
                    return "just";
                default:
                    return "t";
            }
        }
        internal static eTextVerticalType TranslateTextVertical(this string text)
        {
            switch (text)
            {
                case "eaVert":
                    return eTextVerticalType.EastAsianVertical;
                case "mongolianVert":
                    return eTextVerticalType.MongolianVertical;
                case "vert":
                    return eTextVerticalType.Vertical;
                case "vert270":
                    return eTextVerticalType.Vertical270;
                case "wordArtVert":
                    return eTextVerticalType.WordArtVertical;
                case "wordArtVertRtl":
                    return eTextVerticalType.WordArtVerticalRightToLeft;
                default:
                    return eTextVerticalType.Horizontal;
            }
        }
        internal static string TranslateTextVerticalText(this eTextVerticalType value)
        {
            switch (value)
            {
                case eTextVerticalType.EastAsianVertical:
                    return "eaVert";
                case eTextVerticalType.MongolianVertical:
                    return "mongolianVert";
                case eTextVerticalType.Vertical:
                    return "vert";
                case eTextVerticalType.Vertical270:
                    return "vert270";
                case eTextVerticalType.WordArtVertical:
                    return "wordArtVert";
                case eTextVerticalType.WordArtVerticalRightToLeft:
                    return "wordArtVertRtl";
                default:
                    return "horz";
            }
        }
        internal static eStrikeType TranslateStrikeType(this string text)
        {
            switch (text)
            {
                case "dblStrike":
                    return eStrikeType.Double;
                case "sngStrike":
                    return eStrikeType.Single;
                default:
                    return eStrikeType.No;
            }
        }
        internal static string TranslateStrikeTypeText(this eStrikeType value)
        {
            switch (value)
            {
                case eStrikeType.Single:
                    return "sngStrike";
                case eStrikeType.Double:
                    return "dblStrike";
                default:
                    return "noStrike";
            }
        }
        internal static eErrorValueType TranslateErrorValueType(this string text)
        {
            switch (text)
            {
                case "cust":
                    return eErrorValueType.Custom;
                case "fixedVal":
                    return eErrorValueType.FixedValue;
                case "stdDev":
                    return eErrorValueType.StandardDeviation;
                case "stdErr":
                    return eErrorValueType.StandardError;
                default:
                    return eErrorValueType.Percentage;
            }
        }
        internal static string ToEnumString(this eErrorValueType value)
        {
            switch (value)
            {
                case eErrorValueType.Custom:
                    return "cust";
                case eErrorValueType.FixedValue:
                    return "fixedVal";
                case eErrorValueType.StandardDeviation:
                    return "stdDev";
                case eErrorValueType.StandardError:
                    return "stdErr";
                default:
                    return "percentage";
            }
        }
        internal static eSlicerStyle TranslateSlicerStyle(this string value)
        {
            if(string.IsNullOrEmpty(value.Trim()))
            {
                return eSlicerStyle.None;
            }
            else if(value.StartsWith("SlicerStyle", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    return (eSlicerStyle)Enum.Parse(typeof(eSlicerStyle), value.Substring(11));
                }
                catch
                {

                }
            }
            return eSlicerStyle.Custom;
        }
    }
}
