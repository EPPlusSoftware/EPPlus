/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.IO;
using System.Text;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class BorderTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingBorder.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            _pck.Save();
            _pck.Dispose();
            File.Copy(fileName, dirName + "\\DrawingBorderRead.xlsx", true);
        }

        [TestMethod]
        public void BorderFill()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BorderFill");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Border.Fill.Color = Color.Red;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill,shape.Border.Fill.Style);
            Assert.IsNotNull(shape.Border.Fill.SolidFill);
            Assert.AreEqual(Color.Red.ToArgb(), shape.Border.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        [TestMethod]
        public void BorderWidthStyle()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BorderWidthStyle");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Border.Fill.Color = Color.Red;
            shape.Border.Width = 12;
            shape.Border.LineStyle = eLineStyle.Dot;
            shape.Border.CompoundLineStyle = eCompundLineStyle.TripleThinThickThin;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Border.Fill.Style);
            Assert.IsNotNull(shape.Border.Fill.SolidFill);
            Assert.AreEqual(12, shape.Border.Width);
            Assert.AreEqual(eLineStyle.Dot, shape.Border.LineStyle);
            Assert.AreEqual(eCompundLineStyle.TripleThinThickThin, shape.Border.CompoundLineStyle);
        }
        [TestMethod]
        public void BorderAlignRoundJoin()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BorderAlignRoundJoin");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Border.Fill.Color = Color.Red;
            shape.Border.LineStyle = eLineStyle.LongDashDotDot;
            shape.Border.CompoundLineStyle = eCompundLineStyle.Double;
            shape.Border.Alignment = ePenAlignment.Inset;
            shape.Border.LineCap = eLineCap.Square;
            shape.Border.Join = eLineJoin.Round;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Border.Fill.Style);
            Assert.IsNotNull(shape.Border.Fill.SolidFill);
            Assert.AreEqual(eLineStyle.LongDashDotDot, shape.Border.LineStyle);
            Assert.AreEqual(eCompundLineStyle.Double, shape.Border.CompoundLineStyle);
            Assert.AreEqual(ePenAlignment.Inset, shape.Border.Alignment);
            Assert.AreEqual(eLineJoin.Round, shape.Border.Join);
            Assert.AreEqual(eLineCap.Square, shape.Border.LineCap);
        }
        [TestMethod]
        public void BorderMitterJoin()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BorderMiterJoin");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Line);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Border.Fill.Color = Color.Red;
            shape.Border.LineStyle = eLineStyle.LongDashDotDot;
            shape.Border.CompoundLineStyle = eCompundLineStyle.Double;
            shape.Border.LineCap = eLineCap.Flat;
            shape.Border.Join = eLineJoin.Bevel;
            shape.Border.MiterJoinLimit=10000;  //Sets join to Miter

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Border.Fill.Style);
            Assert.IsNotNull(shape.Border.Fill.SolidFill);
            Assert.AreEqual(eLineStyle.LongDashDotDot, shape.Border.LineStyle);
            Assert.AreEqual(eCompundLineStyle.Double, shape.Border.CompoundLineStyle);
            Assert.AreEqual(eLineJoin.Miter, shape.Border.Join);
            Assert.AreEqual(10000, shape.Border.MiterJoinLimit);
            Assert.AreEqual(eLineCap.Flat, shape.Border.LineCap);
        }
        [TestMethod]
        public void BorderEnds()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BorderEnds");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Line);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Border.Fill.Color = Color.Black;
            shape.Border.HeadEnd.Style = eEndStyle.Diamond;
            shape.Border.HeadEnd.Width = eEndSize.Large;
            shape.Border.HeadEnd.Height = eEndSize.Small;
            shape.Border.TailEnd.Style = eEndStyle.Stealth;
            shape.Border.TailEnd.Width = eEndSize.Medium;
            shape.Border.TailEnd.Height = eEndSize.Large;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Border.Fill.Style);
            Assert.IsNotNull(shape.Border.Fill.SolidFill);
            Assert.AreEqual(eEndStyle.Diamond, shape.Border.HeadEnd.Style);
            Assert.AreEqual(eEndSize.Large, shape.Border.HeadEnd.Width);
            Assert.AreEqual(eEndSize.Small, shape.Border.HeadEnd.Height);
            Assert.AreEqual(eEndStyle.Stealth, shape.Border.TailEnd.Style);
            Assert.AreEqual(eEndSize.Medium, shape.Border.TailEnd.Width);
            Assert.AreEqual(eEndSize.Large, shape.Border.TailEnd.Height);
        }

    }
}
