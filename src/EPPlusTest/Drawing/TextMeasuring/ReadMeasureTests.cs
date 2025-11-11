using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.Drawing.TextMeasuring
{
    [TestClass]
    public class ReadMeasureTests: TestBase
    {
        [TestMethod]
        public void ReadShape()
        {
            using(var p = OpenTemplatePackage("ReadText.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var theShape = ws.Drawings[0].As.Shape;

                var height = theShape.TextBody.Paragraphs.GetSizeInPixels(theShape.GetPixelWidth(), theShape.GetPixelHeight(), theShape.Text, theShape.Font);
            }
        }
    }
}
