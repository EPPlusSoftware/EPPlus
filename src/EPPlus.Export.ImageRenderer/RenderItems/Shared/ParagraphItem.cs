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
using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer.RenderItems.Shared
{
    internal abstract class ParagraphItem : RenderItem
    {
        protected List<TextRunItem> TextRuns;

        //ExcelDrawingParagraph Paragraph;
        double RightMargin;
        double LeftMargin;

        eTextAlignment HorizontalAlignment;

        RectBase ParagraphArea;

        ExcelDrawingParagraph Paragraph;

        /// <summary>
        /// X position
        /// </summary>
        protected double XPos;

        /// <summary>
        /// Line-Spacing in pixels
        /// </summary>
        protected double LineSpacing;

        public ParagraphItem(ExcelDrawingParagraph p, RectBase paragraphArea, FontMeasurerTrueType fmExact)
        {
            Paragraph = p;

            //Seperated out in case of some final render item
            //needing to adjust without changing the original values
            RightMargin = p.RightMargin;
            LeftMargin = p.LeftMargin;

            ParagraphArea = paragraphArea;
            HorizontalAlignment = p.HorizontalAlignment;

            LineSpacing = GetParagraphLineSpacingInPixels(fmExact);
            XPos = GetAlignmentHorizontal(HorizontalAlignment);
        }

        private double GetParagraphLineSpacingInPixels(FontMeasurerTrueType fmExact)
        {
            if (Paragraph.LineSpacing.LineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                return Paragraph.LineSpacing.Value.PointToPixel();
            }
            else
            {
                var multiplier = (Paragraph.LineSpacing.Value / 100);
                return multiplier * fmExact.GetSingleLineSpacing().PixelToPoint();
            }
        }

        internal double GetAlignmentHorizontal(eTextAlignment txAlignment)
        {
            var area = ParagraphArea;
            double x = 0;
            switch (txAlignment)
            {
                case eTextAlignment.Left:
                default:
                    x = area.Left;
                    break;
                case eTextAlignment.Center:
                    x = (area.Right / 2) + LeftMargin - RightMargin;
                    break;
                case eTextAlignment.Right:
                    x = area.Right - RightMargin;
                    break;
            }

            return TextUtils.RoundToWhole(x);
        }

        ///// <summary>
        ///// MUST fill the TextRuns list
        ///// </summary>
        //public abstract void InitializeTextRuns<T>() where T : TextRunItem;

        ////internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        ////{
        ////    il = ParagraphArea.Right - RightMargin;
        ////    it = ParagraphArea.Left + LeftMargin;

        ////    ir = ParagraphArea.Top;
        ////    ib = ParagraphArea.Bottom;
        ////}

        //internal RectBase GetBounds()
        //{
        //    GetBounds(out double l, out double t, out double r, out double b);
        //    return new RectBase(l, t, r, b);
        //}
    }
}
