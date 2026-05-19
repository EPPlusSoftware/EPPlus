using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    internal class DrawingTextRun : RenderTextRun
    {
        internal DrawingTextRun(BoundingBox parent, MeasurementFont font, string displayText) : base(parent, font, displayText)
        {
            //Bounds.Name = "TextRun";
            //_currentText = displayText;

            //Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();    
            
            //_measurementFont = font;
            //_isFirstInParagraph = true;

            //FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);
            //Bounds.Height = _measurementFont.Size;
            //if (parent.Height < _measurementFont.Size)
            //{
            //    parent.Height = _measurementFont.Size;
            //}

            ////To get clipping height we need to get the textbody bounds
            //if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            //{
            //    ClippingHeight = parent.Parent.Parent.Position.Y + parent.Parent.Parent.Size.Y;
            //}
            //if (Lines.Count == 1)
            //{
            //    //Bounds.Width = parent.Width;
            //    GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            //}
            //else
            //{
            //    //Measure text.
            //    GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            //}
            //_underLineType = UnderLineType.None;
        }

        internal DrawingTextRun(BoundingBox parent, string text, ExcelTextFont font, string displayText) : base(renderer, parent)
        {
            _originalText = text;

            Bounds.Name = "TextRun";
            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            //_measurer = new FontMeasurerTrueType();            

            _measurementFont = font.GetMeasureFont();
            //_measurer.SetFont(_measurementFont);

            //_fontStyles = _measurementFont.Style;

            _baseline = font.Baseline;
            if (_baseline != 0)
            {
                _measurementFont.Size *= (float)(1 - (Math.Abs(_baseline) / 100));
            }
            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);

            Bounds.Height = _measurementFont.Size;
            if (parent.Height < _measurementFont.Size)
            {
                parent.Height = _measurementFont.Size;
            }
            //_horizontalTextAlignment = TextAlignment.Center;

            if (font.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + font.Fill.Color.To6CharHexString();
            }

            //To get clipping height we need to get the textbody bounds
            if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            {
                ClippingHeight = parent.Parent.Parent.Position.Y + parent.Parent.Parent.Size.Y;
            }
            //if(Lines.Count==1)
            //{
            //    //Bounds.Width = parent.Width;
            //    GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            //}
            //else
            //{
            //    //Measure text.
            //    GetBounds(out double il, out double it, out double ir, out double ib); //TODO: remove when calc works
            //}
            _isItalic = font.Italic;
            _isBold = font.Bold;
            _underLineType = (UnderLineType)font.UnderLine;
            _underlineColor = font.UnderLineColor;
            _strikeType = (StrikeType)font.Strike;
        }

        /// <summary>
        /// If the run has been wrapped more line-breaks may have been added in displayText
        /// </summary>
        /// <param name="run"></param>
        /// <param name="parent"></param>
        /// <param name="displayText"></param>
        internal DrawingTextRun(BoundingBox parent, ExcelParagraphTextRunBase run, string displayText = "") : base(parent)
        {
            _originalText = run.Text;

            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            _measurementFont = run.GetMeasurementFont();

            
            //_fontStyles = _measurementFont.Style;

            _baseline = run.Baseline;
            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);

            Bounds.Height = _measurementFont.Size;

            //_horizontalTextAlignment = run.Paragraph.HorizontalAlignment;

            if (run.Fill.IsEmpty == false && run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }

            //To get clipping height we need to get the textbody bounds
            if( parent!= null && parent.Parent != null && parent.Parent.Parent != null)
            {
               ClippingHeight = ((BoundingBox)parent.Parent.Parent).Bottom;
            }

            if (run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }

            _isItalic = run.FontItalic;
            _isBold = run.FontBold;
            _underLineType = (UnderLineType)run.FontUnderLine;
            _underlineColor = run.UnderLineColor;
            _strikeType = (StrikeType)run.FontStrike;
        }

        ///// <summary>
        ///// Calculates right/bottom
        ///// </summary>
        ///// <param name="il"></param>
        ///// <param name="it"></param>
        ///// <param name="ir"></param>
        ///// <param name="ib"></param>
        //internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        //{
        //    il = Bounds.Left;
        //    it = Bounds.Top;
        //    ir = Bounds.Right;
        //    ib = Bounds.Bottom;
        //}
    }
}
