using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using System;
using System.Globalization;
using System.Text;
using OfficeOpenXml.Style;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextRunItem : TextRunItem
    {
        public SvgTextRunItem(ExcelParagraphTextRunBase run, BoundingBox parent = null, string displayText = "") : base(run, parent, displayText)
        {
        }

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

        public override void Render(StringBuilder sb)
        {
            string finalString = "";

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

                    yIncrease = Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

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
                base.Render(sb);
                //Since final string has been written erase it.
                finalString = "";

                finalString += ">";
                finalString += line;
                finalString += "</tspan>";
            }

            sb.Append(finalString);
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
