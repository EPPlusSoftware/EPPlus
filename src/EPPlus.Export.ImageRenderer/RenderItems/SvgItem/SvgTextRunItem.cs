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
                fontStyleAttributes += "font-style=\"italic\" ";
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
            var xString = $"x =\"{(Bounds.X).ToString(CultureInfo.InvariantCulture)}\" ";

            var currentYEndPos = Bounds.GlobalY;

            for (int i = 0; i < Lines.Count; i++)
            {
                var line = Lines[i];

                finalString += $"<tspan ";
                string visibility = "";

                //Textrun may continue on same line or start a new line
                //Refer to pre-calculated list
                if (YIncreasePerLine[i] != 0)
                {
                    var yIncrease = Fonts.OpenType.Utils.TextUtils.RoundToWhole(YIncreasePerLine[i]);

                    currentYEndPos += yIncrease;
                    if (double.IsNaN(ClippingHeight) == false && currentYEndPos >= ClippingHeight)
                    {
                        visibility = "display=\"none\"";
                    }

                    var yIncreaseString = yIncrease.ToString(CultureInfo.InvariantCulture);
                    var dyString = $"dy=\"{yIncreaseString}px\" ";
                    finalString += dyString;
                    finalString += "x=\"0\" ";
                }
                else
                {
                    finalString += xString;
                }

                finalString += $"{visibility} " + $"{GetFontStyleAttributes()} ";

                if (_measurementFont != null)
                {
                    finalString += $"font-family=\"{_measurementFont.FontFamily},"
                        + $"{_measurementFont.FontFamily}_MSFontService,sans-serif\" "
                        + $"font-size=\"{FontSizeInPixels.ToString(CultureInfo.InvariantCulture)}px\" ";
                }

                sb.Append(finalString);

                //Get color etc.
                //Renders up until this point
                SvgBaseRenderer.BaseRender(sb, this);
                //Since final string has been written in base.render erase it.
                finalString = "";

                finalString += ">";
                finalString += line;
                finalString += "</tspan>";
            }
            sb.Append(finalString);
        }
    }
}
