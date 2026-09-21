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
using System;
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

            SaveAndCleanup(_pck);
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

        [TestMethod]
        public void i2517_BorderAroundShouldNotAffectDimension_Generated()
        {
            var ws = _pck.Workbook.Worksheets.Add("shouldNotAffectDimension");

            var headerRange = ws.Cells["A1:G2"];

            headerRange.Value = "A";
            headerRange.Style.Fill.SetBackground(Color.LightYellow);
            headerRange.Style.Font.Color.SetColor(Color.IndianRed);
            headerRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.MediumDashDot;

            var firstColRange = ws.Cells["A1:A33"];
            firstColRange.Value = "A";

            ws.Cells["B3:G33"].Style.Fill.SetBackground(Color.DarkSeaGreen);
            ws.Cells["B3:G33"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            ws.Cells["B3:G33"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            ws.Cells["B3:G33"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            ws.Cells["B3:G33"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

            firstColRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dotted, false);

            var colCount = ws.Dimension.Columns;
            var titleRange = ws.Cells[1, 1, 1, colCount];

            Assert.AreEqual(7, ws.Dimension.Columns);
            titleRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            Assert.AreEqual(7, ws.Dimension.Columns);


            var bottomRange = ws.Cells[1, 1, 33, 1];
            Assert.AreEqual(33, ws.Dimension.Rows);
            bottomRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, false);
            Assert.AreEqual(33, ws.Dimension.Rows);

            ws.Cells["D34"].Value = "Test";

            Assert.AreEqual(34, ws.Dimension.Rows);
            Assert.AreEqual(7, ws.Dimension.Columns);
            bottomRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, false);
            Assert.AreEqual(34, ws.Dimension.Rows);
            Assert.AreEqual(7, ws.Dimension.Columns);
        }

        [TestMethod]
        public void i2517_BorderAroundShouldNotAffectDimension()
        {
            //This testClass requires the package to exist. Do this to ensure it does even though unused in this test.
            _pck.Workbook.Worksheets.Add("shouldNotAffectDimension_IntentionallyEmpty");

            using (var p = OpenTemplatePackage("i2517.xlsx"))
            {
                var worksheet = p.Workbook.Worksheets[0];
                var colCount = worksheet.Dimension.Columns;
                var titleRange = worksheet.Cells[1, 1, 1, colCount];

                Assert.AreEqual(7, worksheet.Dimension.Columns);
                titleRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Assert.AreEqual(7, worksheet.Dimension.Columns);


                var bottomRange = worksheet.Cells[1, 1, 33, 1];
                Assert.AreEqual(33, worksheet.Dimension.Rows);
                bottomRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                Assert.AreEqual(33, worksheet.Dimension.Rows);
            }
        }

        [TestMethod]
        public void i2525_BorderAroundDropsStylesWhenCalledOnIndividualCellsInAloop()
        {
            var ws = _pck.Workbook.Worksheets.Add("BorderAround_MissingStyles");

            var tb1 = ws.Tables.Add(ws.Cells["B2:K11"],"BorderedTable");
            tb1.TableStyle = OfficeOpenXml.Table.TableStyles.Light15;

            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    //Set borders on every "even" individual cell
                    if (j % 2 > 0 && i % 2 > 0)
                    {
                        ws.Cells[i + 2, j + 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dashed);
                    }
                }
            }
        }
    }
}
