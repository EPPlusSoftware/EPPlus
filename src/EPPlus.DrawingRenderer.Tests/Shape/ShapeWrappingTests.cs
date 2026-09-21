using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.Tests.Shape
{
    [TestClass]
    public class ShapeWrappingTests : TestBase
    {
        [TestMethod]
        public void WrapEveryLetter()
        {
            using(var p = OpenPackage("WrapEveryLetterSvg.xlsx",true))
            {
                var ws = p.Workbook.Worksheets.Add("wrapShapes");

                var txtBox = ws.Drawings.AddTextbox("txtBox1", "MY WORLD");
                txtBox.As.Shape.SetSize(30, 150);

                var svg = txtBox.ToSvg();
                File.WriteAllText(GetOutputFile("svg\\", "WrapEveryLetter.svg").FullName, svg);
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void WrapAndColor()
        {
            using (var p = OpenPackage("WrapInRect.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("wrap");

                var _currentShape = ws.Drawings.AddShape("MyShape", OfficeOpenXml.Drawing.eShapeStyle.Rect);
                _currentShape.SetSize(36, 200);

                _currentShape.Font.Color = System.Drawing.Color.Goldenrod;
                _currentShape.TextBody.LeftInsert = 0;
                _currentShape.TextBody.RightInsert = 0;
                _currentShape.TextBody.TopInsert = 0;
                _currentShape.TextBody.BottomInsert = 0;

                var rt1 = _currentShape.RichText.Add("M", true);
                var rt2 = _currentShape.RichText.Add("Y ", false);
                var rt3 = _currentShape.RichText.Add("W", false);
                var rt4 = _currentShape.RichText.Add("O", false);
                var rt5 = _currentShape.RichText.Add("R", false);
                var rt6 = _currentShape.RichText.Add("L", false);
                var rt7 = _currentShape.RichText.Add("D", false);
                _currentShape.RichText.Add("MY WORLD", true);

                var startColor = KnownColor.Plum;

                _currentShape.RichText.Add("Default world", true);

                foreach (var item in _currentShape.RichText)
                {
                    item.Color = System.Drawing.Color.FromKnownColor(startColor);
                    startColor += 1;
                }

                _currentShape.TextBody.Anchor = eTextAnchoringType.Top;

                var svg = _currentShape.ToSvg();
                File.WriteAllText(GetOutputFile("svg\\", "WrapAndColor.svg").FullName, svg);

                SaveAndCleanup(p);
            }
        }
    }
}
