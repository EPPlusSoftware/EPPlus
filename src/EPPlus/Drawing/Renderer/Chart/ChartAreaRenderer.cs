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
using EPPlus.DrawingRenderer.Svg;
using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAreaRenderer : ChartDrawingObject
    {
        public ChartAreaRenderer(ChartRenderer sc, SvgRenderOptions options) : base(sc)
        {
            if(options.Size.Width.HasValue)
            {
                sc.Bounds.Width = options.Size.WidthPixels;
            }
            if (options.Size.Height.HasValue)
            {
                sc.Bounds.Height = options.Size.HeightPixels;
            }

            Rectangle = new RectRenderItem(sc.Bounds);
        }

        internal override Color? DefaultFillColor { get => ChartRenderer.Theme.ColorScheme.Light1.GetColor(); }
        internal override Color? DefaultBorderColor
        {
            get
            {
                return Color.FromArgb(0x89, 0x89, 0x89);
            }
        }

        internal void InitStyleColors()
        {
            StyleBorderColor1 = GetThemeColorTint(eThemeSchemeColor.Text1, 0.75d);
            StyleBorderColor2 = GetThemeColorTint(eThemeSchemeColor.Background1, 0.75d);
            StyleBorderColor3 = GetThemeColorTint(eThemeSchemeColor.Background1, 0.75d);
            StyleBorderColor4 = GetThemeColorTint(eThemeSchemeColor.Text1, 1d);

            var themedFill = ChartRenderer.Theme.FormatScheme.BorderStyle[0];

            StyleColor1 = GetThemeColorTint(eThemeSchemeColor.Background1, 1d);
            StyleColor2 = GetThemeColorTint(eThemeSchemeColor.Background1, 0.2d);

            //Make this go up by 1 per styleID somehow
            StyleColor3 = GetThemeColorTint(eThemeSchemeColor.Accent1, 1d);

            StyleColor4 = GetThemeColorTint(eThemeSchemeColor.Background1, 0.95d);
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }
    }
}
