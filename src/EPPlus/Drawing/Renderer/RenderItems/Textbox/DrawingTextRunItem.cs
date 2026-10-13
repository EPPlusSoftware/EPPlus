using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;


namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    internal class DrawingTextRunRenderItem : TextRunRenderItem
    {
        /// <summary>
        /// Most basic of all textruns without even a font
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="text"></param>
        /// /// <param name="origRtIndex"></param>
        internal DrawingTextRunRenderItem(BoundingBox parent, string text, int origRtIndex) : base(parent, text, origRtIndex)
        {

        }

        /// <summary>
        /// Imports style info from the text run. The font is taken from the fragment the text was
        /// measured with, never from the model, so measurement and output cannot diverge.
        /// </summary>
        internal void ImportTextRunBase(ExcelParagraphTextRunBase run, IFontFormatBase measuredFont)
        {
            InitializeBase(measuredFont);

            //Text is imported elsewhere Do NOT add the below line back.
            //Empty lines have no text and to replace it with original run.Text
            //is to replace every new line with the entirety of the original textfragment
            //This comment is to be kept as it is easy to believe this function Should import text otherwise
            //_currentText = string.IsNullOrEmpty(_currentText) ? run.Text : _currentText;

            _isFirstInParagraph = run.IsFirstInParagraph;
            _baseline = run.Baseline;
            ImportExcelStyleInfo(run.Fill, run.FontItalic, run.FontBold, run.FontUnderLine, run.UnderLineColor, run.FontStrike);
            SetClippingHeightToCurrentTextBoxBottom(Parent);
        }

        /// <summary>
        /// Text font holds style info.
        /// baseFont is most likely a OpenTypeFontInfoBase made out of the font but we don't want to 'new' it every time we import
        /// </summary>
        /// <param name="font"></param>
        /// <param name="baseFont"></param>
        internal void ImportExcelTextFont(ExcelTextFont font, IFontFormatBase baseFont)
        {
            InitializeBase(baseFont);
            _baseline = font.Baseline;

            //Adjusts visual font size for sub and superscript
            if (_baseline != 0)
            {
                _measurementFont.Size *= (float)(1 - (Math.Abs(_baseline) / 100));
            }
            //Must be done after font adjustment
            //Parent and clipping height must be calculated dependent on content
            ImportExcelStyleInfo(font.Fill, font.Italic, font.Bold, font.UnderLine, font.UnderLineColor, font.Strike);
            //Assumes texbox uses auto-size?
            AdjustParentAndSetClippingHeight((BoundingBox)Parent);
        }

        private void ImportExcelStyleInfo(ExcelDrawingFill fill, bool italic, bool bold, eUnderLineType uType, Color uColor, eStrikeType strikeType)
        {
            ImportDrawingFill(fill);
            ImportRichTextInfo(italic, bold, uType, uColor, strikeType);
        }

        private void AdjustParentAndSetClippingHeight(BoundingBox parent)
        {
            if (parent.Height < _measurementFont.Size)
            {
                parent.Height = _measurementFont.Size;
            }

            CalculateClippingHeightFromTextBodyParent();
        }

        void ImportDrawingFill(ExcelDrawingFill fill)
        {
            if (fill.IsEmpty == false && fill.Style == eFillStyle.SolidFill)
            {
                Style.FillColor = "#" + fill.Color.To6CharHexString();
            }

            //Backup? Should probably be removed or fallback
            if (fill.Style == eFillStyle.SolidFill)
            {
                Style.FillColor = "#" + fill.Color.To6CharHexString();
            }
        }

        void ImportRichTextInfo(bool italic, bool bold, eUnderLineType uType, Color uColor, eStrikeType strikeType)
        {
            _isItalic = italic;
            _isBold = bold;
            _underLineType = (eDrawingUnderLineType)uType;
            _underlineColor = uColor;
            _strikeType = (eDrawingStrikeType)strikeType;
        }

        void SetClippingHeightToCurrentTextBoxBottom(Transform parent)
        {
            //To get clipping height we need to get the textbody bounds
            if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            {
                ClippingHeight = ((BoundingBox)parent.Parent.Parent).Bottom;
            }
        }

        public override RenderItem Clone()
        {
            throw new NotImplementedException();
        }
    }
}
