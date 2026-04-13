using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.Constants;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils;
using System.Drawing;
using System.Globalization;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;


namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    internal class LinearGradient : RenderItem
    {
        internal DrawGradientFill GradientFillExtra;
        internal double Degrees;

        string _id;
        bool userSpaceOnUse = false;

        public LinearGradient(DrawingBase renderer, string id) : base(renderer)
        {
            _id = id;
        }

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<linearGradient id=\"{_id}\" y1=\"1\" x2=\"0\" gradientUnits=\"userSpaceOnUse\">");
            SetStopColors(sb, GradientFillExtra, PathFillMode.Norm);
            sb.Append("</linearGradient>");
        }

        private string GetOpacity(ExcelDrawingColorManager c)
        {
            var opacityTransform = c.Transforms?.FirstOrDefault(x => x.Type == OfficeOpenXml.Drawing.Style.Coloring.eColorTransformType.Alpha);
            if (opacityTransform == null) return "";

            return $"stop-opacity=\"{opacityTransform.Value.ToString("0")}%\"";
        }

        private void SetStopColors(StringBuilder defSb, DrawGradientFill gradientFill, PathFillMode fillMode)
        {
            int ix = 0;

            //Svg requires starting at 0 and moving towards 100% Excel sometimes starts at 100
            //Sort to get around that
            var sortedGradientColors = gradientFill.Colors.OrderBy(x => x.Position);

            foreach (var c in sortedGradientColors)
            {
                var color = ColorUtils.GetAdjustedColor(fillMode, c.Color);
                // TODO: check if ix should be increased...?
                defSb.Append($"<stop offset=\"{c.Position}%\" stop-color=\"#{color.To6CharHexString()}\" />");
            }
        }
    }
}
