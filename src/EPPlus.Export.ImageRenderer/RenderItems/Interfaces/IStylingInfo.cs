using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    public interface StylingInfo
    {
        string FillColor { get; set; }
        string FilterName { get; set; }
        //public DrawGradientFill GradientFill { get; set; }
        //SvgFillType FillType { get; set; }
        double? FillOpacity { get; set; }
        string BorderColor { get; set; }
        //DrawGradientFill BorderGradientFill { get; set; }
        ExcelDrawingPatternFill PatternFill { get; }
        ExcelDrawingBlipFill BlipFill { get; }
        double? BorderWidth { get; set; }
        double[] BorderDashArray { get; set; }
        int StrokeMiterLimit { get; set; }
        eCompoundLineStyle CompoundLineStyle { get; set; }
        double? BorderDashOffset { get; set; }
        eLineCap LineCap { get; set; }
        //SvgLineJoin LineJoin { get; set; }
        double? BorderOpacity { get; set; }
        PathFillMode FillColorSource { get; set; }
        PathFillMode BorderColorSource { get; set; }
        double? GlowRadius { get; }
        string GlowColor { get; }
        ExcelDrawingOuterShadowEffect OuterShadowEffect { get; }
    }
}
