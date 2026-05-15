using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.Fill;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    public interface StylingInfo
    {
        IBorder Border { get; set; }
        IFill Fill { get; set; }
        string FilterName { get; set; }
        //public RenderGradientFill GradientFill { get; set; }
        //FillType FillType { get; set; }
        //RenderGradientFill BorderGradientFill { get; set; }
        ExcelDrawingPatternFill PatternFill { get; }
        ExcelDrawingBlipFill BlipFill { get; }
        int StrokeMiterLimit { get; set; }
        eCompoundLineStyle CompoundLineStyle { get; set; }
        eLineCap LineCap { get; set; }
        //SvgLineJoin LineJoin { get; set; }
        double? GlowRadius { get; }
        string GlowColor { get; }
        ExcelDrawingOuterShadowEffect OuterShadowEffect { get; }
    }
}
