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
        ///  TextRunBase holds style info
        ///  baseFont is most likely a OpenTypeFontInfoBase made out of the font but we don't want to 'new' it every time we import
        /// </summary>
        /// <param name="run"></param>
        /// <param name="baseFont"></param>
        internal void ImportTextRunBase(ExcelParagraphTextRunBase run, IFontFormatBase baseFont)
        {
            InitializeBase(new FontFormatBase(run.GetMeasureFont()));
            _currentText = string.IsNullOrEmpty(_currentText) ? run.Text : _currentText;
            _isFirstInParagraph = run.IsFirstInParagraph;
            _baseline = run.Baseline;
            ImportExcelStyleInfo(run.Fill, run.FontItalic, run.FontBold, run.FontUnderLine, run.UnderLineColor, run.FontStrike);
            SetClippingHeightToCurrentTextBoxBottom((BoundingBox)Bounds.Parent);
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
            AdjustParentAndSetClippingHeight((BoundingBox)Bounds.Parent);
        }

        /// <summary>
        /// Import textrun with only font data
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="font"></param>
        /// <param name="displayText"></param>
        /// <param name="adjustParent"></param>
        internal DrawingTextRunRenderItem(BoundingBox parent, IFontFormatBase font, string displayText, bool adjustParent = true) : base(parent, font, displayText)
        {
            //Parent and clipping height must be calculated dependent on content
            if (adjustParent)
            {
                AdjustParentAndSetClippingHeight(parent);
            }
            else
            {
                //If auto-size is not on?
                SetClippingHeightToCurrentTextBoxBottom(parent);
            }
            //Since only font there is no direct style info to import
        }

        /// <summary>
        /// Import TextRun from Default paragraph properties/ExcelTextFont
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="text"></param>
        /// <param name="font">Legacy format</param>
        /// <param name="displayText"></param>
        internal DrawingTextRunRenderItem(BoundingBox parent, string text, ExcelTextFont font, string displayText) : base(parent, text, new FontFormatBase(font.GetMeasureFont()), displayText)
        {
            _baseline = font.Baseline;

            //Adjusts visual font size for sub and superscript
            if (_baseline != 0)
            {
                _measurementFont.Size *= (float)(1 - (Math.Abs(_baseline) / 100));
            }
            //Must be done after font adjustment
            //Parent and clipping height must be calculated dependent on content
            ImportExcelStyleInfo(font.Fill, font.Italic, font.Bold, font.UnderLine, font.UnderLineColor, font.Strike);
            AdjustParentAndSetClippingHeight(parent);
        }

        /// <summary>
        /// Import text run from ParagraphTextRun
        /// </summary>
        /// <param name="run">new format</param>
        /// <param name="parent"></param>
        /// <param name="displayText"></param>
        internal DrawingTextRunRenderItem(BoundingBox parent, ExcelParagraphTextRunBase run, string displayText = "") : base(parent, run.Text, new FontFormatBase(run.GetMeasurementFont()), displayText)
        {
            //This is pre-determined/irrelevant here and does not need to be calculated as sizes are already what they should
            _isFirstInParagraph = false; 
            //Has getXmlNodePercentage, therefore no need for conversion
            _baseline = run.Baseline;

            ImportExcelStyleInfo(run.Fill, run.FontItalic, run.FontBold, run.FontUnderLine, run.UnderLineColor, run.FontStrike);

            //This one cannot change parent size as it is pre-determined. No autosize etc. therefore no AdjustParentAndSetClippingHeight and different clipping height.

            SetClippingHeightToCurrentTextBoxBottom(parent);
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
                FillColor = "#" + fill.Color.To6CharHexString();
            }

            //Backup? Should probably be removed or fallback
            if (fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + fill.Color.To6CharHexString();
            }
        }

        void ImportRichTextInfo(bool italic, bool bold, eUnderLineType uType, Color uColor, eStrikeType strikeType)
        {
            _isItalic = italic;
            _isBold = bold;
            _underLineType = (UnderLineType)uType;
            _underlineColor = uColor;
            _strikeType = (StrikeType)strikeType;
        }

        void SetClippingHeightToCurrentTextBoxBottom(BoundingBox parent)
        {
            //To get clipping height we need to get the textbody bounds
            if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            {
                ClippingHeight = ((BoundingBox)parent.Parent.Parent).Bottom;
            }
        }
    }
}
