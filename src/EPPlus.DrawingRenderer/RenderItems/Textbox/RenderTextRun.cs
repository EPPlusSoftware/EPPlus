using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Globalization;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class RenderTextRun
    {
        public RenderTextRun(BoundingBox parent, MeasurementFont font, string displayText)
        {
        }

        public RenderTextRun(BoundingBox parent, ExcelParagraphTextRunBase run, string displayText = "")
        {
        }

        public RenderTextRun(BoundingBox parent, string text, ExcelTextFont font, string displayText)
        {
                
        }
        string GetFontStyleAttributes()
        {
            string fontStyleAttributes = " ";

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

                fontStyleAttributes += "text-decoration=\" ";
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
            var xString = $"x =\"{(Bounds.Left.PointToPixelString())}\" ";

            var currentYEndPos = Bounds.Position.Y; // Global position Y
            finalString += $"<tspan ";
            string visibility = "";

            double fontSize=FontSizeInPixels;
            if(_baseline!=0)
            { 
                finalString += $" dy=\"{(fontSize.PointToPixel()*-_baseline/100D).ToString(CultureInfo.InvariantCulture)}px\" ";  //For sub/superscript, move the text up/down by baseline% of font size. Negative value moves up, positive moves down.
            }
            finalString += xString;
            var yString = $" y=\"{YPosition.PointToPixelString()}px\" ";
            finalString += yString;

            currentYEndPos += YPosition;

            if (double.IsNaN(ClippingHeight) == false && currentYEndPos >= ClippingHeight)
            {
                //visibility = " display=\"none\"";
            }
            finalString += visibility;
            finalString += $"{GetFontStyleAttributes()}";

            if (_measurementFont != null)
            {
                finalString += $"font-family=\"{_measurementFont.FontFamily},"
                    + $"{_measurementFont.FontFamily}_MSFontService,sans-serif\" "
                    + $"font-size=\"{fontSize.ToString(CultureInfo.InvariantCulture)}px\" ";
            }

            sb.Append(finalString);
            //Get color etc.
            //Renders up until this point
            SvgBaseRenderer.BaseRender(sb, this);
            //Since final string has been written in base.render erase it.
            finalString = "";

            finalString += ">";
            finalString += _currentText;
            finalString += "</tspan>";
            
            sb.Append(finalString);
        }
    }
}
