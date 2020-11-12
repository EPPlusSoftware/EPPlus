using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Globalization;
using System.Linq;
namespace OfficeOpenXml.Drawing.Controls
{
    public class ExcelControlMargin
    {
        private ExcelControl _control;
        private XmlHelper _vmlHelper;
        string[] _suffixes;
        internal ExcelControlMargin(ExcelControl control)
        {
            _control = control;
            _vmlHelper = XmlHelperFactory.Create(control._vmlProp.NameSpaceManager, control._vmlProp.TopNode.ParentNode);
            _suffixes=((eMeasurementUnits[])Enum.GetValues(typeof(eMeasurementUnits))).Where(x=>x!=eMeasurementUnits.Pixels).Select(x=>x.TranslateString()).ToArray();
            
            Automatic = _vmlHelper.GetXmlNodeString("@o:insetmode") == "auto";
            var margin = _vmlHelper.GetXmlNodeString("v:textbox/@inset");

            var v = margin.GetCsvPosition(0);
            LeftMargin = GetMarginValue(v);
            LeftMarginUnit = GetMarginUnit(v);

            v = margin.GetCsvPosition(1);
            TopMargin = GetMarginValue(v);
            TopMarginUnit = GetMarginUnit(v);

            v = margin.GetCsvPosition(2);
            RightMargin = GetMarginValue(v);
            RightMarginUnit = GetMarginUnit(v);

            v = margin.GetCsvPosition(3);
            BottomMargin = GetMarginValue(v);
            BottomMarginUnit = GetMarginUnit(v);
        }
        private double GetMarginValue(string v)
        {
            if (string.IsNullOrEmpty(v)) return 0;
            if(_suffixes.Any(x => v.EndsWith(x)))
            {
                return ConvertUtil.GetValueDouble(v.Substring(0, v.Length - 2));
            }
            return ConvertUtil.GetValueDouble(v);
        }
        private eMeasurementUnits GetMarginUnit(string v)
        {
            foreach(eMeasurementUnits u in Enum.GetValues(typeof(eMeasurementUnits)))
            {
                if (v.EndsWith(u.TranslateString()))
                {
                    return u;
                }
            }
            return eMeasurementUnits.Pixels;
        }
        /// <summary>
        /// Sets the unit of measurement for all margins.
        /// </summary>
        /// <param name="unit"></param>
        public void SetUnit(eMeasurementUnits unit)
        {
            LeftMarginUnit = unit;
            TopMarginUnit = unit;
            RightMarginUnit = unit;
            BottomMarginUnit = unit;
        }

        internal void UpdateXml()
        {
            if (Automatic)
            {
                _vmlHelper.SetXmlNodeString("@o:insetmode", "auto");
            }
            else
            {
                _vmlHelper.DeleteNode("@o:insetmode");    //Custom
            }

            if (LeftMargin != 0 && TopMargin != 0 && RightMargin != 0 && BottomMargin != 0)
            {
                var v =
                    GetStringMargin(LeftMargin, LeftMarginUnit) + "," +
                    GetStringMargin(TopMargin, TopMarginUnit) + "," +
                    GetStringMargin(RightMargin, RightMarginUnit) + "," +
                    GetStringMargin(BottomMargin, BottomMarginUnit);

                _vmlHelper.SetXmlNodeString("v:textbox/@inset", v);
            }
            else
            {
                _vmlHelper.DeleteNode("v:textbox/@inset");
            }
        }

        private string GetStringMargin(double leftMargin, eMeasurementUnits unit)
        {
            return leftMargin.ToString(CultureInfo.InvariantCulture) + unit.TranslateString();
        }

        public bool Automatic
        {
            get;
            set;
        }
        public double LeftMargin 
        {
            get;
            set;
        }
        public eMeasurementUnits LeftMarginUnit
        {
            get;
            set;
        }
        public double RightMargin
        {
            get;
            set;
        }
        public eMeasurementUnits RightMarginUnit
        {
            get;
            set;
        }
        public double TopMargin
        {
            get;
            set;
        }
        public eMeasurementUnits TopMarginUnit
        {
            get;
            set;
        }
        public double BottomMargin
        {
            get;
            set;
        }
        public eMeasurementUnits BottomMarginUnit
        {
            get;
            set;
        }
        private string GetWithSuffixMeasure(double value, eMeasurementUnits unit)
        {
            var v = value.ToString(CultureInfo.InvariantCulture)+unit.ToEnumString();
            return v + unit.TranslateString();
        }

        private double CovertToPt(string v)
        {
            int ratio;
            if (v.EndsWith("mm"))
            {
                v = v.Substring(0, v.Length - 2);
                ratio = ExcelDrawing.EMU_PER_MM;
            }
            else if (v.EndsWith("cm"))
            {
                v = v.Substring(0, v.Length - 2);
                ratio = ExcelDrawing.EMU_PER_CM;
            }
            else if (v.EndsWith("pt"))
            {
                v = v.Substring(0, v.Length - 2);
                ratio = ExcelDrawing.EMU_PER_POINT;
            }
            else if (v.EndsWith("pc"))
            {
                v = v.Substring(0, v.Length - 2);
                ratio = ExcelDrawing.EMU_PER_PICA;
            }
            else
            {
                ratio = ExcelDrawing.EMU_PER_PIXEL;
            }

            return ConvertUtil.GetValueDouble(v) * ratio / ExcelDrawing.EMU_PER_POINT;
        }
    }
}
