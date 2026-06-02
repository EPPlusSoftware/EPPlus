using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    internal class SvgTextRunRenderer : SvgBaseRenderer<TextRunRenderItem>
    {
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
            if (textRun._underLineType != eDrawingUnderLineType.None | textRun._strikeType != eDrawingStrikeType.No)
            {
                string content = "";
                if (textRun._underLineType != eDrawingUnderLineType.None)
                {
                    switch (textRun._underLineType)
                    {
                        case eDrawingUnderLineType.Single:
                            content += "underline";
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
                            content += "underline";
                            break;
                            //throw new NotImplementedException("Not implemented yet");
                    }
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
                    fontStyleAttributes += "text-decoration=\" ";
                    fontStyleAttributes += content;
                    fontStyleAttributes += "\" ";
                }
            }
            return fontStyleAttributes;
        }

        public override void Render(TextRunRenderItem textRun)
        {
            string finalString = "";
            var xString = $"x =\"{(textRun.Bounds.Left.PointToPixelString())}\" ";

            var currentYEndPos = textRun.Bounds.Position.Y; // Global position Y
            finalString += $"<tspan ";
            string visibility = "";

            double fontSize = textRun.FontSizeInPixels;
            if (textRun._baseline != 0)
            {
                finalString += $" dy=\"{(fontSize.PointToPixel() * -textRun._baseline / 100D).ToString(CultureInfo.InvariantCulture)}px\" ";  //For sub/superscript, move the text up/down by baseline% of font size. Negative value moves up, positive moves down.
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

            var sb = OutputStream;
            sb.Append(finalString);
            //Get color etc.
            //Renders up until this point
            RenderBase(textRun);
            //Since final string has been written in base.render erase it.
            finalString = "";

            finalString += ">";
            finalString += textRun._currentText;
            finalString += "</tspan>";

            sb.Append(finalString);
        }
    }
}
