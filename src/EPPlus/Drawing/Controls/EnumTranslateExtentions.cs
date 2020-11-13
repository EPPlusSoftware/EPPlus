using OfficeOpenXml.Utils.Extentions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Controls
{
    public static class EnumTranslateExtentions
    {
        internal static eMeasurementUnits TranslatePresetColor(this string v)
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
                default:
                    return eMeasurementUnits.Pixels;
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
        internal static eShapeOrienation TranslateShapeOrienation(this string v)
        {
            switch (v)
            {
                case "top-to-bottom":
                    return eShapeOrienation.TopToBottom;
                case "bottom-to-top":
                    return eShapeOrienation.BottomToTop;
                default:
                    return eShapeOrienation.Auto;
            }
        }
        internal static string TranslateString(this eShapeOrienation v)
        {
            switch (v)
            {
                case eShapeOrienation.TopToBottom:
                    return "top-to-bottom";
                case eShapeOrienation.BottomToTop:
                    return "bottom-to-top";
                default:
                    return "auto";
            }
        }
        
    }
}
