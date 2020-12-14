using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.Drawing.Vml
{
    internal static class VmlConvertUtil
    {
        internal static double GetOpacityFromStringVml(string v)
        {
            if (string.IsNullOrEmpty(v))
            {
                return 0;
            }
            if (v.EndsWith("f", StringComparison.OrdinalIgnoreCase))
            {
                ConvertUtil.TryParseNumericString(v.Substring(0, v.Length - 1), out double d);
                return (d / 0x10000) * 100;
            }
            else if (v.EndsWith("%"))
            {
                ConvertUtil.TryParseNumericString(v.Substring(0, v.Length - 1), out double d);
                return d;
            }
            else
            {
                ConvertUtil.TryParseNumericString(v.Substring(0, v.Length - 1), out double d);
                return d * 100;
            }
        }
        internal static double ConvertToEMU(double v, eMeasurementUnits measure)
        {
            int ratio;
            switch (measure)
            {
                case eMeasurementUnits.Millimeters:
                    ratio = ExcelDrawing.EMU_PER_MM;
                    break;
                case eMeasurementUnits.Centimeters:
                    ratio = ExcelDrawing.EMU_PER_CM;
                    break;
                case eMeasurementUnits.Points:
                    ratio = ExcelDrawing.EMU_PER_POINT;
                    break;
                case eMeasurementUnits.Picas:
                    ratio = ExcelDrawing.EMU_PER_PICA;
                    break;
                case eMeasurementUnits.Inches:
                    ratio = ExcelDrawing.EMU_PER_US_INCH;
                    break;
                default:
                    ratio = ExcelDrawing.EMU_PER_PIXEL;
                    break;
            }

            return v * ratio;
        }

        //internal static double GetValueInPoints(string s)
        //{
        //    double ix = s.Length, convert;
        //    if (s.EndsWith("pt", StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        ix = s.Length - 2;
        //        convert = ExcelDrawing.EMU_PER_POINT;
        //    }
        //    else if (s.EndsWith("mm", StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        ix = s.Length - 2;
        //        convert = ExcelDrawing.EMU_PER_MM;
        //    }
        //    else if (s.EndsWith("mm", StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        ix = s.Length - 2;
        //        convert = ExcelDrawing.EMU_PER_CM;
        //    }
        //    else if (s.EndsWith("pc", StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        ix = s.Length - 2;
        //        convert = ExcelDrawing.EMU_PER_PICA;
        //    }
        //    else if (s.EndsWith("in", StringComparison.InvariantCultureIgnoreCase))
        //    {
        //        ix = s.Length - 2;
        //        convert = ExcelDrawing.EMU_PER_POINT;
        //    }
        //}
    }
}
