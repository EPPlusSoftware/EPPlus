using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class GenericFill
    {
        internal double _degree;
        internal double _right;
        internal double _bottom;
        internal bool _isLinear = false;
        internal bool _isSolid = false;
        internal bool _isGradient = false;
        internal ExcelFillStyle _patternType;

        internal GenericColor _color1;
        internal GenericColor _color2;

        //public GenericFill() { }

        public GenericFill(ExcelFillXml fill) 
        {
            _isSolid = fill.PatternType == ExcelFillStyle.Solid;
            _patternType = fill.PatternType;

            if (fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                _isGradient = true;

                _degree = gf.Degree;
                _right = gf.Right;
                _bottom = gf.Bottom;
                _isLinear = gf.Type == ExcelFillGradientType.Linear;

                _color1 = new GenericColor(gf.GradientColor1);
                _color2 = new GenericColor(gf.GradientColor2);
            }
            else
            {
                _color1 = new GenericColor(fill.BackgroundColor);
                _color2 = new GenericColor(fill._patternColor);
            }

        }

        public GenericFill(ExcelDxfFill fill) 
        {
            _isSolid = fill.PatternType == ExcelFillStyle.Solid;
            _patternType = fill.PatternType.Value;

            if (fill.Gradient.HasValue)
            {
                _isGradient = true;

                _degree = fill.Gradient.Degree.Value;
                _right = fill.Gradient.Right.Value;
                _bottom = fill.Gradient.Bottom.Value;
                _isLinear = fill.Gradient.GradientType == eDxfGradientFillType.Linear;

                _color1 = new GenericColor(fill.Gradient.Colors[0].Color);
                _color2 = new GenericColor(fill.Gradient.Colors[1].Color);
            }
            else
            {
                _color1 = new GenericColor(fill.BackgroundColor);
                _color2 = new GenericColor(fill.PatternColor);
            }
        }


    }
}
