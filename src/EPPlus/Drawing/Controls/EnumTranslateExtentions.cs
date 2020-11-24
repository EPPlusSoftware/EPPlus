/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extentions;

namespace OfficeOpenXml.Drawing.Controls
{
    public static class EnumTranslateExtentions
    {
        internal static eMeasurementUnits TranslateMeasurementUnits(this string v)
        {
            switch(v)
            {
                case "mm":
                    return eMeasurementUnits.Millimeters;
                case "cm":
                    return eMeasurementUnits.Centimeters;
                case "pt":
                    return eMeasurementUnits.Points;
                case "in":
                    return eMeasurementUnits.Inches;
                case "pc":
                    return eMeasurementUnits.Picas;
                case "px":
                    return eMeasurementUnits.Pixels;
                default:
                    return eMeasurementUnits.EMUs;
            }
        }
        internal static string TranslateString(this eMeasurementUnits v)
        {
            switch (v)
            {
                case eMeasurementUnits.Millimeters:
                    return "mm";
                case eMeasurementUnits.Centimeters:
                    return "cm";
                case eMeasurementUnits.Points:
                    return "pt";
                case eMeasurementUnits.Inches:
                    return "in";
                case eMeasurementUnits.Picas:
                    return "pc";
                case eMeasurementUnits.Pixels:
                    return "px";
                default:
                    return "";  //Blank is Pixels, px
            }
        }

        internal static eLayoutFlow TranslateLayoutFlow(this string v)
        {
            switch (v)
            {
                case "horizontal-ideographic":
                    return eLayoutFlow.HorizontalIdeographic;
                case "vertical-ideographic":
                    return eLayoutFlow.VerticalIdeographic;
                default:
                    return v.ToEnum(eLayoutFlow.Horizontal);
            }
        }
        internal static string TranslateString(this eLayoutFlow v)
        {
            switch (v)
            {
                case eLayoutFlow.HorizontalIdeographic:
                    return "horizontal-ideographic";
                case eLayoutFlow.VerticalIdeographic:
                    return "vertical-ideographic";
                default:
                    return v.ToString().ToLower();  
            }
        }
        internal static eShapeOrientation TranslateShapeOrientation(this string v)
        {
            switch (v)
            {
                case "top-to-bottom":
                    return eShapeOrientation.TopToBottom;
                case "bottom-to-top":
                    return eShapeOrientation.BottomToTop;
                default:
                    return eShapeOrientation.Auto;
            }
        }
        internal static string TranslateString(this eShapeOrientation v)
        {
            switch (v)
            {
                case eShapeOrientation.TopToBottom:
                    return "top-to-bottom";
                case eShapeOrientation.BottomToTop:
                    return "bottom-to-top";
                default:
                    return "auto";
            }
        }
        
    }
}
