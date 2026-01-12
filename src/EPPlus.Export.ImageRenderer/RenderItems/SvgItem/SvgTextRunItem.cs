using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextRunItem : TextRunRenderItem
    {
        public SvgTextRunItem(ExcelParagraphTextRunBase run, BoundingBox parent = null, string displayText = "") : base(run, parent, displayText)
        {
        }

        //public override void Render(StringBuilder sb)
        //{
        //    var runElement = new SvgElement("tspan");
        //    runElement.AddAttribute("x", Bounds.X);
        //    runElement.AddAttribute("y", Bounds.Y + FontSizeInPixels);
        //    runElement.AddAttribute("font-size", $"{FontSizeInPixels}px");

        //    runElement.Content = _currentText;
        //    var retStr = RenderSvgElement(runElement);
        //    sb.AppendLine(retStr);
        //    sb.AppendLine("</tspan>");
        //}

        string GetFontStyleAttributes()
        {
            string fontStyleAttributes = "";

            if (_isItalic)
            {
                fontStyleAttributes += " font-style=\"italic\" ";
            }
            if (_isBold)
            {
                fontStyleAttributes += "font-weight=\"bold\" ";
            }
            if (_underLineType != eUnderLineType.None | _strikeType != eStrikeType.No)
            {

                fontStyleAttributes += "text-decoration=\"";
                if (_underLineType != eUnderLineType.None)
                {
                    switch (_underLineType)
                    {
                        case eUnderLineType.Single:
                            fontStyleAttributes += "underline";
                            break;
                        //These are all css only apparently
                        //case eUnderLineType.Double:
                        //    fontStyleAttributes += "double";
                        //    break;
                        //case eUnderLineType.Dotted:
                        //    fontStyleAttributes += "dotted";
                        //    break;
                        //case eUnderLineType.Dash:
                        //    fontStyleAttributes += "dashed";
                        //    break;
                        //case eUnderLineType.Wavy:
                        //    fontStyleAttributes += "wavy";
                        //    break;
                        default:
                            fontStyleAttributes += "underline";
                            break;
                            //throw new NotImplementedException("Not implemented yet");
                    }
                }

                if (_strikeType == eStrikeType.Single)
                {
                    //Has to check if Both underline and strike
                    if (_underLineType != eUnderLineType.None)
                    {
                        fontStyleAttributes += ",";
                    }
                    fontStyleAttributes += "line-through";
                }

                fontStyleAttributes += "\" ";
            }

            return fontStyleAttributes;
        }

        public void BaseRender(StringBuilder sb)
        {
            if (string.IsNullOrEmpty(FillColor) == false)
            {
                sb.Append($"fill=\"{FillColor}\" ");
                if (FillOpacity != null && FillOpacity != 1)
                {
                    sb.Append($"opacity=\"{FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
                }
            }
            if (string.IsNullOrEmpty(FilterName) == false)
            {
                sb.Append($"filter=\"{FilterName}\" ");
            }

            if (BorderWidth.HasValue && string.IsNullOrEmpty(BorderColor) == false)
            {
                sb.Append($"stroke=\"{BorderColor}\" ");
            }
            if (BorderWidth.HasValue)
            {
                var v = BorderWidth.Value * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL;
                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

                if (BorderDashArray != null)
                {
                    var BorderDashArrayStr = BorderDashArray.Select(x =>
                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
                }
            }

            sb.Append($"stroke-miterlimit =\"8\"");
        }

        public override void Render(StringBuilder sb)
        {
            string finalString = "";
            //bool useBaselineSpacing = double.IsNaN(BaselineSpacing) == false;
            //Lines = SplitIntoLines(currentText);

            foreach (var line in Lines)
            {
                finalString += $"<tspan ";
                string visibility = "";
                //Despite new textrun it could still be on the same line as previous textrun
                //Therefore only do line increase if we are first in paragraph or if we are not Lines[0].
                //This as line == Lines[0] && isFirstInParagraph == false means we are continuing on the same line as previous textRun
                //This is important if for example we have rich text where two letters on the same line has different colors.
                if (line != Lines[0] | _isFirstInParagraph)
                {
                    var yIncrease = _isFirstInParagraph ? BaseLineSpacing : LineSpacingPerNewLine;
                    _isFirstInParagraph = false;

                    yIncrease = EPPlus.Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

                    //_yEndPos += yIncrease;
                    //if (Double.IsNaN(ClippingHeight) == false && _yEndPos >= ClippingHeight)
                    //{
                    //    visibility = "display=\"none\"";
                    //}

                    var yIncreaseString = yIncrease.ToString(CultureInfo.InvariantCulture);
                    var xString = $"x =\"{(Bounds.X).ToString(CultureInfo.InvariantCulture)}\" ";
                    var dyString = $"dy =\"{yIncreaseString}px\" ";
                    finalString += xString;
                    finalString += dyString;
                }

                finalString += $"{visibility} " + $"{GetFontStyleAttributes()} ";
                if (_measurementFont != null)
                {
                    finalString += $" font-family=\"{_measurementFont.FontFamily},"
                        + $"{_measurementFont.FontFamily}_MSFontService,sans-serif\" "
                        + $"font-size=\"{_measurementFont.Size.ToString(CultureInfo.InvariantCulture)}px\" ";
                }
                sb.Append(finalString);

                //Get color etc.
                //Renders up until this point
                BaseRender(sb);
                //Since final string has been written erase it.
                finalString = "";

                finalString += ">";
                finalString += line;
                finalString += "</tspan>";
            }

            sb.Append(finalString);
            //////throw new NotImplementedException();
            //foreach (var line in Lines)
            //{
            //    sb.AppendLine($"<tspan x=\"{Bounds.X}\" y=\"{Bounds.Y + FontSizeInPixels}\" font-size=\"{FontSizeInPixels}px\" >");
            //    sb.AppendLine(line);
            //    sb.AppendLine("</tspan>");
            //}
        }
    }
}
