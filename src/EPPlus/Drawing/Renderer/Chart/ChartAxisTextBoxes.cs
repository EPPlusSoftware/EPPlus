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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Collections.Generic;
using System.Drawing;
using EPPlus.Graphics;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAxisTextBoxes : ChartDrawingObject
    {
        internal string AxisName = "";

        internal override Color? DefaultFillColor { get; }
        internal ChartAxisRenderer _axis;
        internal ChartAxisTextBoxes(ChartAxisRenderer axis) : base(axis.ChartRenderer)
        {
            _axis = axis;
            DefaultFillColor = Color.Transparent;
        }

        internal List<DrawingTextBox> TextBoxes
        {
            get;
            set;
        }=new List<DrawingTextBox>();


        public override void AppendRenderItems(List<Transform> renderItems)
        {
            if (TextBoxes != null && TextBoxes.Count > 0)
            {
                var axisTxtBoxGroup = new GroupRenderItem(ChartRenderer.Bounds);
                axisTxtBoxGroup.Name = AxisName;
                axisTxtBoxGroup.Top = _axis.Rectangle.Top;
                axisTxtBoxGroup.Left = _axis.Rectangle.Left;
                _axis.Rectangle.Top = 0;
                _axis.Rectangle.Left = 0;
                foreach (var tb in TextBoxes)
                {
                    tb.AppendRenderItems(axisTxtBoxGroup.ChildObjects);
                }
                renderItems.Add(axisTxtBoxGroup);
            }

        }
        internal override Color? DefaultBorderColor => null;

        internal override Color? GetDefaultFillColor()
        {
            return DefaultFillColor;
        }

        internal override Color? GetDefaultBorderColor()
        {
            return DefaultBorderColor;
        }
    }
}