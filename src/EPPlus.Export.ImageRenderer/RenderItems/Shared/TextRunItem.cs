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
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using EPPlus.Fonts.OpenType.Utils;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusImageRenderer.RenderItems.Shared
{
    internal abstract class TextRunItem
    {
        internal RectBase BoundingBox = new RectBase();

        internal Coordinate origin = new Coordinate(0,0);
        internal ExcelParagraphTextRunBase TextRun;

        double defaultFontSize;

        string FontName = "";

        //Unnecesary??
        double TextLengthInPixels;

        List<string> Lines;

        const float minValueParagraphSize = int.MinValue;
        MeasurementFont measurementFont;

        internal TextRunItem(ExcelParagraphTextRunBase textRun)
        {
            TextRun = textRun;
            Lines = SplitIntoLines();

            measurementFont = textRun.GetMeasurementFont();

            FontName = measurementFont.FontFamily;

            //double fontSize = TextRun.FontSize < 0 || TextRun.FontSize == minValueParagraphSize ? defaultFontSize : TextRun.FontSize;
            
            double fontSize = measurementFont.Size;
        }
        ///// <summary>
        ///// Construct instance
        ///// </summary>
        ///// <typeparam name="T"></typeparam>
        ///// <returns></returns>
        //public abstract T Create<T>(ExcelDrawing drawing, ExcelParagraphTextRunBase textRun) where T : TextRunItem;

        //private void GetDefaultFontNameAndSize(ExcelFont? nsFont, ExcelShape shape, out string fontName, out double fontSize)
        //{
        //    fontName = string.IsNullOrEmpty(shape.Font.LatinFont) ? shape.Font.ComplexFont : shape.Font.LatinFont;

        //    fontSize = shape.Font.Size;
        //    if (string.IsNullOrEmpty(fontName)) fontName = nsFont?.Name ?? _theme.FontScheme.MajorFont.First().Typeface;
        //    if (fontSize <= 0 && nsFont != null) fontSize = nsFont.Size;
        //}

        //public override SvgItemType Type => SvgItemType.Rect;

        internal void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = origin.X;
            it = origin.Y;

            ir = CalculateRightPositionInPixels();
            ib = CalculateBottomPositionInPixels();

            BoundingBox.Left = il;
            BoundingBox.Top = it;
            BoundingBox.Right = ir;
            BoundingBox.Bottom = ib;
        }

        internal double CalculateBottomPositionInPixels()
        {
            var bottomPosition = (((double)TextRun.FontSize).PointToPixel() + origin.Y) * GetNumberOfLines();
            return bottomPosition;
        }

        internal double CalculateTextWidth(string targetString)
        {
            var textMesurer = new FontMeasurerTrueType(measurementFont);
            textMesurer.MeasureWrappedTextCells = true;
            var width = textMesurer.MeasureTextWidth(targetString);

            return width;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = origin.X;

            if (numLines <= 0 || string.IsNullOrEmpty(TextRun.Text))
            {
                TextLengthInPixels = 0;
                return retPos;
            }

            double longestWidth = -1;

            for (int i = 0; i < numLines; i++)
            {
                var text = Lines[i];
                var width = CalculateTextWidth(Lines[i]);

                if(width > longestWidth)
                {
                    longestWidth = width;
                }
            }

            TextLengthInPixels = longestWidth;

            retPos += longestWidth;

            return retPos;
        }

        internal double GetNumberOfLines()
        {
            if(TextRun != null)
            {
                SplitIntoLines();
                return Lines.Count;
            }
            else
            {
                return 0;
            }
        }

        internal List<string> SplitIntoLines()
        {
            return TextRun.Text.Split(new string[]{ "\r\n"}, System.StringSplitOptions.None).ToList();
        }

        //private string GetTextRunFontName(ExcelParagraphTextRunBase textRun)
        //{
        //    if (string.IsNullOrEmpty(textRun.LatinFont))
        //    {
        //        if (string.IsNullOrEmpty(textRun.ComplexFont))
        //        {
        //            return defaultFontName;
        //        }
        //        else
        //        {
        //            return textRun.ComplexFont;
        //        }
        //    }
        //    else
        //    {
        //        return textRun.LatinFont;
        //    }
        //}

        ////internal double CalculateRightPositionInPixels()
        //{

        //}

        //internal Coordinate SetPosition(double x, double y)
        //{
        //    if (origin == null)
        //    {
        //        origin = new Coordinate(x, y);
        //    }
        //    else
        //    {
        //        origin.X = x; origin.Y = y;
        //    }
        //    return origin;
        //}

        //internal void SetPositionX(double x)
        //{
        //    if (origin == null)
        //    {
        //        origin = new Coordinate(0, 0);
        //    }
        //    origin.X = x;
        //}

        //internal void SetPositionY(double y)
        //{
        //    if (origin == null)
        //    {
        //        origin = new Coordinate(0, 0);
        //    }
        //    origin.Y = y;
        //}

        //public override void Render(StringBuilder sb)
        //{
        //    string finalString = "";

        //    yPosition += TextUtils.RoundToWhole(SingleLineSpacingInPixels);
        //    string visibility = "";
        //    if (yPosition >= yClipHeight)
        //    {
        //        visibility = "display=\"none\"";
        //    }
        //    var xIndent = GetAlignmentHorizontal(horizontalTextAlignment);
        //    var horizAttr = GetHorizontalAlignmentAttribute(xIndent);

        //    finalString += $"<tspan {visibility} {horizAttr} font-size=\"{(fontSizeInPixels / 16).ToString(CultureInfo.InvariantCulture)}em\" dy=\"{dy.ToString(CultureInfo.InvariantCulture)}px\">";
        //    finalString += textContent;
        //    finalString += "</tspan>";

        //    sb.Append(finalString);
        //    //throw new NotImplementedException();
        //}

        //private string GetHorizontalAlignmentAttribute(double indentX)
        //{
        //    string ret = "";
        //    var xStr = indentX.ToString(CultureInfo.InvariantCulture);

        //    switch (horizontalTextAlignment)
        //    {
        //        default:
        //        case eTextAlignment.Left:
        //            ret = $"text-anchor=\"start\" x=\"{xStr}\" ";
        //            break;
        //        case eTextAlignment.Center:
        //            ret = $"text-anchor=\"middle\" x=\"{xStr}\" ";
        //            break;
        //        case eTextAlignment.Right:

        //            ret = $"text-anchor=\"end\" x=\"{xStr}\" ";
        //            break;
        //    }

        //    return ret;
        //}

        //internal override RenderItem Clone(SvgShape svgDocument)
        //{
        //    throw new NotImplementedException();
        //}

        //internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        //{
        //    throw new NotImplementedException();
        //}

        //internal double GetAlignmentHorizontal(eTextAlignment txAlignment)
        //{
        //    var area = TextBox.GetTextArea();
        //    double x = 0;
        //    switch (txAlignment)
        //    {
        //        case eTextAlignment.Left:
        //        default:
        //            x = area.Left;
        //            break;
        //        case eTextAlignment.Center:
        //            x = (area.Right / 2) + TextBox.Margin.Left;
        //            break;
        //        case eTextAlignment.Right:
        //            x = area.Right;
        //            break;
        //    }

        //    return TextUtils.RoundToWhole(x);
        //}
    }
}
