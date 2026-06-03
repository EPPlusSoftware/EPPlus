using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    internal class SvgTextRunRenderer : SvgBaseRenderer<TextRunRenderItem>
    {
        string UnderlineColorString = string.Empty;

        public SvgTextRunRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }
        string GetFontStyleAttributes(TextRunRenderItem textRun)
        {
            string fontStyleAttributes = " ";

            if (textRun._isItalic)
            {
                fontStyleAttributes += "font-style=\"italic\" ";
            }
            if (textRun._isBold)
            {
                fontStyleAttributes += "font-weight=\"bold\" ";
            }

            string content = "";

            if (textRun._underLineType != eDrawingUnderLineType.None | textRun._strikeType != eDrawingStrikeType.No)
            {
                string underlineContent = "";
                if (textRun._underLineType != eDrawingUnderLineType.None)
                {
                    switch (textRun._underLineType)
                    {
                        case eDrawingUnderLineType.Single:
                            underlineContent += "underline";
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
                            underlineContent += "underline";
                            break;
                            //throw new NotImplementedException("Not implemented yet");
                    }
                    if(textRun._underlineColor != System.Drawing.Color.Empty)
                    {
                        UnderlineColorString = $"<tspan style=\"fill: #{textRun._underlineColor.To6CharHexString()}; {{0}} \">{{1}}</tspan>";
                    }
                    content += underlineContent;
                    //Underline color
                    //We can change the underline color using double tspans
                    // <text>SVG with a <tspan style="fill: red; text-decoration: underline;"><tspan style="fill:black;">colored underline</tspan></tspan>section.</text>
                }

                if (textRun._strikeType == eDrawingStrikeType.Single)
                {
                    //Has to check if Both underline and strike
                    if (textRun._underLineType != eDrawingUnderLineType.None)
                    {
                        content += ",";
                    }
                    content += "line-through";
                }

                if(string.IsNullOrEmpty(content) == false)
                {
                    if (string.IsNullOrEmpty(UnderlineColorString))
                    {
                        fontStyleAttributes += $" text-decoration=\"{content}\" ";
                    }
                    else
                    {
                        UnderlineColorString = string.Format(UnderlineColorString, $"text-decoration: {content};", "{0}");
                    }
                }
            }
            return fontStyleAttributes;
        }

        public override void Render(TextRunRenderItem textRun)
        {
            var sbStartidx = OutputStream.Length -1;

            string finalString = "";
            var xString = $"x =\"{(textRun.Bounds.Left.PointToPixelString())}\" ";

            var currentYEndPos = textRun.Bounds.Position.Y; // Global position Y
            finalString += $"<tspan ";
            string visibility = "";

            double fontSize = textRun.FontSizeInPixels;
            //var baseLine = -textRun._baseline / 100D;
            if (textRun._baseline != 0)
            {
                //For sub/superscript, move the text up/down by baseline% of font size. Negative value moves up, positive moves down.
                finalString += $" dy=\"{(fontSize.PointToPixel() * -textRun._baseline / 100D).ToString(CultureInfo.InvariantCulture)}px\" ";
                //fontSize *= (1 - Math.Abs(baseLine));
            }
            finalString += xString;
            var yString = $" y=\"{textRun.YPosition.PointToPixelString()}px\" ";
            finalString += yString;

            currentYEndPos += textRun.YPosition;

            if (double.IsNaN(textRun.ClippingHeight) == false && currentYEndPos >= textRun.ClippingHeight)
            {
                //visibility = " display=\"none\"";
            }
            finalString += visibility;
            finalString += $"{GetFontStyleAttributes(textRun)}";


            if (textRun._measurementFont != null)
            {
                finalString += $"font-family=\"{textRun._measurementFont.Family},"
                    + $"{textRun._measurementFont.Family}_MSFontService,sans-serif\" "
                    + $"font-size=\"{fontSize.ToString(CultureInfo.InvariantCulture)}px\" ";
            }

            if(string.IsNullOrEmpty(textRun.FillColor) == false)
            {
                finalString += $" style=\"fill: {textRun.FillColor};\" ";
            }

            //Avoid rendering fill color as we do so via style
            var temp = textRun.FillColor;
            textRun.FillColor = null;

            var sb = new StringBuilder();
            sb.Append(finalString);

            //Get color etc.
            //Renders up until this point (must be done to end the attribute addings so that text content can then be added)
            RenderBaseToSpecified(textRun, sb);

            textRun.FillColor = temp;

            //Since final string has been written in base.render erase it.
            finalString = "";

            finalString += ">";
            //Add actual text content
            finalString += textRun._currentText;
            finalString += "</tspan>";
            sb.Append(finalString);

            var textRunString = sb.ToString();

            //Wrap in another tspan to apply underline color if neccesary
            if(string.IsNullOrEmpty(UnderlineColorString) == false)
            {
                textRunString = string.Format(UnderlineColorString, textRunString);
            }

            OutputStream.Append(textRunString);
            UnderlineColorString = string.Empty;
        }
    }
}
