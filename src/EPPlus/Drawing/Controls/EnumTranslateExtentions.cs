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
    }
}
