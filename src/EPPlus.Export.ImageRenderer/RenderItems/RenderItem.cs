/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Graphics;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Drawing;
using System.Text;
using EPPlusColorConverter = OfficeOpenXml.Utils.TypeConversion.ColorConverter;
namespace EPPlusImageRenderer.RenderItems
{
    internal abstract class RenderItem : RenderItemBase
    {
        internal protected DrawingBase DrawingRenderer { get; }
        internal RenderItem(DrawingBase renderer)
        {
            DrawingRenderer = renderer;
        }

        internal RenderItem(DrawingBase renderer, BoundingBox parent)
        {
            Bounds.Parent = parent;
            DrawingRenderer = renderer; 
        }
        //internal abstract void GetBounds(out double il, out double it, out double ir, out double ib);
        internal virtual void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left;
            it = Bounds.Top;
            ir = Bounds.Right;
            ib = Bounds.Bottom;
        }

        //internal bool IsEndOfGroup { get; set; } = false;
        public string FillColor { get; set; }
        public string FilterName { get; set; }
        public DrawGradientFill GradientFill { get; set; }
        public SvgFillType FillType { get; set; }
        public double? FillOpacity { get; set; }
        public string BorderColor { get; set; }
        public DrawGradientFill BorderGradientFill { get; set; }
        public ExcelDrawingPatternFill PatternFill { get; private set; }
        public ExcelDrawingBlipFill BlipFill { get; private set; }
        public double? BorderWidth { get; set; }
        public double[] BorderDashArray { get; set; }
        public double? BorderDashOffset { get; set; }
        public eLineCap LineCap { get; set; } = eLineCap.Flat;
        public SvgLineJoin LineJoin { get; set; } = SvgLineJoin.Miter;
        public double? BorderOpacity { get; set; }
        public PathFillMode FillColorSource { get; set; } = PathFillMode.Norm;
        public PathFillMode BorderColorSource { get; set; } = PathFillMode.Norm;

        protected void CloneBase(RenderItem item)
        {
            item.FillColor = FillColor;
            item.FillOpacity = FillOpacity;
            item.BorderWidth = BorderWidth;
            item.BorderColor = BorderColor;
            item.BorderDashArray = BorderDashArray;
            item.BorderDashOffset = BorderDashOffset;
            item.BorderOpacity = BorderOpacity;
            item.LineJoin = LineJoin;
            item.LineCap = LineCap;
            item.FillColorSource = FillColorSource;
        }

        internal virtual void SetDrawingPropertiesFill(ExcelDrawingFill fill, ExcelDrawingColorManager color)
        {
            switch (fill.Style)
            {

                case eFillStyle.PatternFill:
                    PatternFill = fill.PatternFill;
                    break;
                case eFillStyle.BlipFill:
                    BlipFill = fill.BlipFill;
                    break;
                default:
                    SetDrawingPropertiesFill((ExcelDrawingFillBasic)fill, color);
                    break;
            }
        }
        internal virtual void SetDrawingPropertiesFill(ExcelDrawingFillBasic fill, ExcelDrawingColorManager color)
        {
            switch (fill.Style)
            {
                case eFillStyle.NoFill:
                    if (fill.IsEmpty)
                    {
                        FillColor = GetFillColor(fill, color, FillColorSource);
                    }
                    else
                    {
                        FillColor = "none";
                    }
                    break;
                case eFillStyle.SolidFill:
                    FillColor = GetFillColor(fill, color, FillColorSource);
                    break;
                case eFillStyle.GradientFill:
                    GradientFill = new DrawGradientFill(DrawingRenderer.Theme, fill.GradientFill);
                    FillColor = null;
                    break;
            }
        }
        internal virtual void SetDrawingPropertiesBorder(ExcelDrawingBorder border, ExcelChartStyleColorManager color, bool hasBorder, double defaultWidth=1.5)
        {
            switch (border.Fill.Style)
            {
                case eFillStyle.NoFill:
                    if(border.Fill.IsEmpty)
                    {
                        BorderColor = GetFillColor(border.Fill, color, BorderColorSource);
                    }
                    else
                    {
                        BorderColor = "none";
                    }
                    break;
                case eFillStyle.SolidFill:
                    BorderColor = GetFillColor(border.Fill, color, BorderColorSource);
                    BorderGradientFill = null;
                    break;
                case eFillStyle.GradientFill:
                    BorderGradientFill = new DrawGradientFill(DrawingRenderer.Theme, border.Fill.GradientFill);
                    BorderColor = null;
                    break;
            }

            if (hasBorder && BorderColorSource != PathFillMode.None)
            {
                BorderWidth = border.Width == 0 ? defaultWidth : border.Width;
                if(border.LineStyle.HasValue && border.LineStyle!=eLineStyle.Solid)
                {
                    BorderDashArray = GetDashArray(border);
                }
                if(border.CompoundLineStyle!=eCompundLineStyle.Single)
                {
                    //TODO:Add support double compound borders.
                }
            }
        }
        private double[] GetDashArray(ExcelDrawingBorder border)
        {
            var lw = (int)Math.Round(border.Width * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL);
            switch (border.LineStyle)
            {
                case eLineStyle.Dot:
                    return new double[]{ lw, 4 * lw };
                case eLineStyle.DashDot:
                    return new double[] { 4 * lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.Dash:
                    return new double[] { 4 * lw, 3 * lw };
                case eLineStyle.LongDash:
                    return new double[] { 8 * lw, 3 * lw };
                case eLineStyle.LongDashDot:
                    return new double[] { 8 * lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.LongDashDotDot:
                    return new double[] { 8 * lw, 3 * lw, lw, 3 * lw, lw, 3 * lw };
                case eLineStyle.SystemDash:
                    return new double[] { 3 * lw, lw };
                case eLineStyle.SystemDot:
                    return new double[] { lw, lw };
                case eLineStyle.SystemDashDot:
                    return new double[] { 3 * lw, lw, lw, lw };
                case eLineStyle.SystemDashDotDot:
                    return new double[] { 3 * lw, lw, lw, lw, lw, lw };
            }
            return null;
        }

        private string GetFillColor(ExcelDrawingFillBasic fill, ExcelDrawingColorManager styleFillColor, PathFillMode fillColorSource)
        {
            if (fillColorSource == PathFillMode.None)
            {
                return "none";
            }

            Color fc;
            if (fill == null || fill.Style == eFillStyle.NoFill)
            {
                if (styleFillColor == null)
                {
                    fc = EPPlusColorConverter.GetThemeColor(DrawingRenderer.Theme.ColorScheme.Accent1);
                }
                else
                {
                    fc = EPPlusColorConverter.GetThemeColor(DrawingRenderer.Theme, styleFillColor);
                }
            }
            else if (fill.Style == eFillStyle.SolidFill)
            {
                fc = EPPlusColorConverter.GetThemeColor(DrawingRenderer.Theme, fill.SolidFill.Color);
            }
            else
            {
                return string.Empty;
            }

            fc = ColorUtils.GetAdjustedColor(fillColorSource, fc);
            return "#" + fc.ToArgb().ToString("x8").Substring(2);
        }

        //internal void SetTheme(ExcelTheme theme)
        //{
        //    _theme = theme;
        //}
    }
    /// <summary>
    /// Base class for any item rendered.
    /// </summary>
    internal abstract class RenderItemBase
    {
        internal BoundingBox Bounds = new BoundingBox();
        public abstract RenderItemType Type { get; }
        public abstract void Render(StringBuilder sb);
    }
}