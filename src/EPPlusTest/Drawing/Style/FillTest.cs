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
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Style;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class FillTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingFill.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            File.Copy(fileName, dirName + "\\DrawingFillRead.xlsx", true);
        }
        #region SolidFill
        [TestMethod]
        public void ColorProperty()
        {
            //Setup
            var expected = Color.Blue;
            var ws = _pck.Workbook.Worksheets.Add("SolidFill");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = expected;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.Rgb, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
            Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        [TestMethod]
        public void SolidFill_NoColor()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("SolidFillNoColor");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.None, shape.Fill.SolidFill.Color.ColorType);
        }

        [TestMethod]
        public void SolidFill_Color()
        {
            //Setup
            var expected = Color.Green;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromSolidFill");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetRgbColor(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.Rgb, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
            Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        [TestMethod]
        public void NoFill()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("NoFill");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.NoFill;

            //Assert
            Assert.AreEqual(eFillStyle.NoFill, shape.Fill.Style);
        }
        [TestMethod]
        public void SolidFill_ColorPreset()
        {
            //Setup
            var expected = ePresetColor.Red;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromPresetClr");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetPresetColor(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.PresetColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.PresetColor.Color);
        }
        [TestMethod]
        public void SolidFill_ColorScheme()
        {
            //Setup
            var expected = eSchemeColor.Accent6;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromSchemeClr");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.Scheme, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.SchemeColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SchemeColor.Color);
        }
        [TestMethod]
        public void SolidFill_ColorPercentage()
        {
            //Setup
            var expectedR = 51;
            var expectedG = 49;
            var expectedB = 50;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromColorPrc");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetRgbPercentageColor(expectedR, expectedG, expectedB);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.RgbPercentage, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbPercentageColor);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.RgbPercentageColor.RedPercentage);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.RgbPercentageColor.GreenPercentage);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.RgbPercentageColor.BluePercentage);
        }
        [TestMethod]
        public void SolidFill_ColorHsl()
        {
            //Setup
            var expectedHue = 180;
            var expectedLum = 15;
            var expectedSat = 50;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromColorHcl");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetHslColor(expectedHue, expectedSat, expectedLum);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.Hsl, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.HslColor);
            Assert.AreEqual(expectedHue, shape.Fill.SolidFill.Color.HslColor.Hue);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.HslColor.Luminance);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.HslColor.Saturation);
        }
        [TestMethod]
        public void SolidFill_ColorSystem()
        {
            //Setup
            var expected = eSystemColor.Background;
            var ws = _pck.Workbook.Worksheets.Add("SolidFillFromColorSystem");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetSystemColor(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eDrawingColorType.System, shape.Fill.SolidFill.Color.ColorType);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.SystemColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SystemColor.Color);
        }
        #endregion
        #region Transform
        [TestMethod]
        public void Transparancy()
        {
            //Setup
            var expected = 45;
            var ws = _pck.Workbook.Worksheets.Add("Transparancy");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Red;
            shape.Fill.Transparancy = expected;

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(expected, shape.Fill.Transparancy);
            Assert.AreEqual(100 - expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void TransformAlpha()
        {
            //Setup
            var expected = 45;
            var ws = _pck.Workbook.Worksheets.Add("Alpha");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Red;
            shape.Fill.SolidFill.Color.Transforms.AddAlpha(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(100 - expected, shape.Fill.Transparancy);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void TransformTint()
        {
            //Setup
            var expected = 30;
            var ws = _pck.Workbook.Worksheets.Add("Tint");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Blue;
            shape.Fill.SolidFill.Color.Transforms.AddTint(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.Tint, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void TransformShade()
        {
            //Setup
            var expected = 95;
            var ws = _pck.Workbook.Worksheets.Add("Shade");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Blue;
            shape.Fill.SolidFill.Color.Transforms.AddShade(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.Shade, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void TransformInverse_true()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Inverse_set");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Blue;
            shape.Fill.SolidFill.Color.Transforms.AddInverse();

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.Inv, shape.Fill.SolidFill.Color.Transforms[0].Type);
        }
        [TestMethod]
        public void TransformAlphaModulation()
        {
            //Setup
            var expected = 50;
            var ws = _pck.Workbook.Worksheets.Add("AlphaModulation");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Blue;
            shape.Fill.SolidFill.Color.Transforms.AddAlpha(20);
            shape.Fill.SolidFill.Color.Transforms.AddAlphaModulation(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.AlphaMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void TransformAlphaOffset()
        {
            //Setup
            var expected = -10;
            var ws = _pck.Workbook.Worksheets.Add("AlphaOffset");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Blue;
            shape.Fill.SolidFill.Color.Transforms.AddAlpha(20);
            shape.Fill.SolidFill.Color.Transforms.AddAlphaOffset(expected);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.AlphaOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void TransformColorPercentage()
        {
            //Setup
            var expectedR = 30;
            var expectedG = 60;
            var expectedB = 20;
            var ws = _pck.Workbook.Worksheets.Add("TransColorPerc");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Black;
            shape.Fill.SolidFill.Color.Transforms.AddRed(expectedR);
            shape.Fill.SolidFill.Color.Transforms.AddGreen(expectedG);
            shape.Fill.SolidFill.Color.Transforms.AddBlue(expectedB);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.Red, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.Green, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
            Assert.AreEqual(eColorTransformType.Blue, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void TransformColorModulation()
        {
            //Setup
            var expectedR = 3.33;
            var expectedG = 50;
            var expectedB = 25600;
            var ws = _pck.Workbook.Worksheets.Add("TransColorMod");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Gray;
            shape.Fill.SolidFill.Color.Transforms.AddRedModulation(expectedR);
            shape.Fill.SolidFill.Color.Transforms.AddGreenModulation(expectedG);
            shape.Fill.SolidFill.Color.Transforms.AddBlueModulation(expectedB);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.RedMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.GreenMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.BlueMod, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void TransformColoOffset()
        {
            //Setup
            var expectedR = 10;
            var expectedG = -20;
            var expectedB = 30;
            var ws = _pck.Workbook.Worksheets.Add("TransColorOffset");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Gray;
            shape.Fill.SolidFill.Color.Transforms.AddRedOffset(expectedR);
            shape.Fill.SolidFill.Color.Transforms.AddGreenOffset(expectedG);
            shape.Fill.SolidFill.Color.Transforms.AddBlueOffset(expectedB);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.RedOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.GreenOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.BlueOff, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void TransformHslOffset()
        {
            //Setup
            var expectedLum = 10;
            var expectedSat = -20;
            var ws = _pck.Workbook.Worksheets.Add("TransHslOffset");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Gray;
            shape.Fill.SolidFill.Color.Transforms.AddLuminanceOffset(expectedLum);
            shape.Fill.SolidFill.Color.Transforms.AddSaturationOffset(expectedSat);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.LumOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.SatOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void TransformHslModulation()
        {
            //Setup
            var expectedLum = 50;
            var expectedSat = 200;
            var ws = _pck.Workbook.Worksheets.Add("TransHslModulation");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Color = Color.Gray;
            shape.Fill.SolidFill.Color.Transforms.AddLuminanceModulation(expectedLum);
            shape.Fill.SolidFill.Color.Transforms.AddSaturationModulation(expectedSat);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsInstanceOfType(shape.Fill.SolidFill.Color.RgbColor, typeof(ExcelDrawingRgbColor));
            Assert.AreEqual(eColorTransformType.LumMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.SatMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        #endregion
        #region Gradiant
        [TestMethod]
        public void Gradient()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Gradient");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;
            shape.Fill.GradientFill.Colors.AddRgb(0, Color.Red);
            shape.Fill.GradientFill.Colors.AddRgb(50.35, Color.Yellow);
            shape.Fill.GradientFill.Colors.AddRgb(100, Color.Blue);
            shape.Fill.GradientFill.RotateWithShape = true;

            shape.Fill.GradientFill.TileFlip = eTileFlipMode.None;
            Assert.AreEqual(eTileFlipMode.None, shape.Fill.GradientFill.TileFlip);
            shape.Fill.GradientFill.TileFlip = eTileFlipMode.X;
            Assert.AreEqual(eTileFlipMode.X, shape.Fill.GradientFill.TileFlip);
            shape.Fill.GradientFill.TileFlip = eTileFlipMode.Y;
            Assert.AreEqual(eTileFlipMode.Y, shape.Fill.GradientFill.TileFlip);
            shape.Fill.GradientFill.TileFlip = eTileFlipMode.XY;
            Assert.AreEqual(eTileFlipMode.XY, shape.Fill.GradientFill.TileFlip);

            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(true, shape.Fill.GradientFill.RotateWithShape);
        }
        [TestMethod]
        public void GradientNotSet()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GradientNotSet");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;

            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.AreEqual(0, shape.Fill.GradientFill.Colors.Count);
            Assert.AreEqual(false, shape.Fill.GradientFill.RotateWithShape);
        }
        [TestMethod]
        public void GradientCircularPath()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GradientCircular");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;
            shape.Fill.GradientFill.Colors.AddRgb(0, Color.Green);
            shape.Fill.GradientFill.Colors.AddRgb(50.35, Color.Olive);
            shape.Fill.GradientFill.Colors.AddRgb(100, Color.Gray);
            shape.Fill.GradientFill.ShadePath = eShadePath.Circle;

            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(eShadePath.Circle, shape.Fill.GradientFill.ShadePath);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
        }
        [TestMethod]
        public void GradientRectPath()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GradientRect");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;
            shape.Fill.GradientFill.Colors.AddRgb(0, Color.White);
            shape.Fill.GradientFill.Colors.AddRgb(50.35, Color.Gray);
            shape.Fill.GradientFill.Colors.AddRgb(100, Color.Black);
            shape.Fill.GradientFill.ShadePath = eShadePath.Rectangle;
            shape.Fill.GradientFill.FocusPoint.BottomOffset = 20;
            shape.Fill.GradientFill.FocusPoint.LeftOffset = 20;

            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(eShadePath.Rectangle, shape.Fill.GradientFill.ShadePath);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
            Assert.AreEqual(20, shape.Fill.GradientFill.FocusPoint.BottomOffset);
            Assert.AreEqual(20, shape.Fill.GradientFill.FocusPoint.LeftOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
        }
        [TestMethod]
        public void GradientShapePath()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GradientShape");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Heart);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;
            shape.Fill.GradientFill.Colors.AddRgb(0, Color.LightBlue);
            shape.Fill.GradientFill.Colors.AddRgb(40, Color.Blue);
            shape.Fill.GradientFill.Colors.AddRgb(100, Color.DarkBlue);
            shape.Fill.GradientFill.ShadePath = eShadePath.Shape;

            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(eShadePath.Shape, shape.Fill.GradientFill.ShadePath);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
        }
        [TestMethod]
        public void Gradient_AddMethods()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GradientAddMethods");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.GradientFill;
            shape.Fill.GradientFill.Colors.AddRgb(0, Color.Red);
            shape.Fill.GradientFill.Colors.AddRgbPercentage(22.55, 40, 50, 60.5);
            shape.Fill.GradientFill.Colors.AddHsl(37.42, 180, 50, 60);
            shape.Fill.GradientFill.Colors.AddPreset(55.2, ePresetColor.BlueViolet);
            shape.Fill.GradientFill.Colors.AddScheme(66.2, eSchemeColor.Background2);
            shape.Fill.GradientFill.Colors.AddSystem(88.2, eSystemColor.GradientActiveCaption);


            //Assert
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(eDrawingColorType.RgbPercentage, shape.Fill.GradientFill.Colors[22.55].Color.ColorType); //Verify index for position

            //RGB
            Assert.AreEqual(0, shape.Fill.GradientFill.Colors[0].Position);
            Assert.AreEqual(eDrawingColorType.Rgb, shape.Fill.GradientFill.Colors[0].Color.ColorType);
            Assert.AreEqual(Color.Red.ToArgb(), shape.Fill.GradientFill.Colors[0].Color.RgbColor.Color.ToArgb());

            //RGB Percent
            Assert.AreEqual(22.55, shape.Fill.GradientFill.Colors[1].Position);
            Assert.AreEqual(eDrawingColorType.RgbPercentage, shape.Fill.GradientFill.Colors[1].Color.ColorType);
            Assert.AreEqual(40, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.RedPercentage);
            Assert.AreEqual(50, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.GreenPercentage);
            Assert.AreEqual(60.5, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.BluePercentage);

            //Hsl Percent
            Assert.AreEqual(37.42, shape.Fill.GradientFill.Colors[2].Position);
            Assert.AreEqual(eDrawingColorType.Hsl, shape.Fill.GradientFill.Colors[2].Color.ColorType);
            Assert.AreEqual(180, shape.Fill.GradientFill.Colors[2].Color.HslColor.Hue);
            Assert.AreEqual(50, shape.Fill.GradientFill.Colors[2].Color.HslColor.Saturation);
            Assert.AreEqual(60, shape.Fill.GradientFill.Colors[2].Color.HslColor.Luminance);

            //Preset
            Assert.AreEqual(55.2, shape.Fill.GradientFill.Colors[3].Position);
            Assert.AreEqual(eDrawingColorType.Preset, shape.Fill.GradientFill.Colors[3].Color.ColorType);
            Assert.AreEqual(ePresetColor.BlueViolet, shape.Fill.GradientFill.Colors[3].Color.PresetColor.Color);

            //Scheme color
            Assert.AreEqual(66.2, shape.Fill.GradientFill.Colors[4].Position);
            Assert.AreEqual(eDrawingColorType.Scheme, shape.Fill.GradientFill.Colors[4].Color.ColorType);
            Assert.AreEqual(eSchemeColor.Background2, shape.Fill.GradientFill.Colors[4].Color.SchemeColor.Color);

            //Scheme color
            Assert.AreEqual(88.2, shape.Fill.GradientFill.Colors[5].Position);
            Assert.AreEqual(eDrawingColorType.System, shape.Fill.GradientFill.Colors[5].Color.ColorType);
            Assert.AreEqual(eSystemColor.GradientActiveCaption, shape.Fill.GradientFill.Colors[5].Color.SystemColor.Color);
        }
        #endregion
        [TestMethod]
        public void PatternDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("PatternDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.PatternFill;
            shape.Fill.PatternFill.BackgroundColor.SetRgbColor(Color.Red);
            shape.Fill.PatternFill.ForegroundColor.SetRgbColor(Color.Blue);

            //Assert
            Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.AreEqual(eFillPatternStyle.Pct5, shape.Fill.PatternFill.PatternType);
        }
        [TestMethod]
        public void PatternDefaultColorCheck()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("PatternDefaultColorCheck");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.PatternFill;

            //Assert
            Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.AreEqual(eDrawingColorType.Scheme, shape.Fill.PatternFill.BackgroundColor.ColorType);
            Assert.AreEqual(eSchemeColor.Background1, shape.Fill.PatternFill.BackgroundColor.SchemeColor.Color);
            Assert.AreEqual(eDrawingColorType.Scheme, shape.Fill.PatternFill.ForegroundColor.ColorType);
            Assert.AreEqual(eSchemeColor.Text1, shape.Fill.PatternFill.ForegroundColor.SchemeColor.Color);
            Assert.AreEqual(eFillPatternStyle.Pct5, shape.Fill.PatternFill.PatternType);
        }
        [TestMethod]
        public void PatternCross()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("PatternCross");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.PatternFill;
            shape.Fill.PatternFill.PatternType = eFillPatternStyle.Cross;
            shape.Fill.PatternFill.BackgroundColor.SetSchemeColor(eSchemeColor.Accent4);
            shape.Fill.PatternFill.ForegroundColor.SetSchemeColor(eSchemeColor.Background2);

            //Assert
            Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.AreEqual(eFillPatternStyle.Cross, shape.Fill.PatternFill.PatternType);
        }
        #region BlipFill
        [TestMethod]
        public void BlipFill_DefaultSettings()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipFill");

            var shape = AddBlip(ws, 1, "Shape1", false, 0, 0);

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
            Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
        }
        [TestMethod]
        public void BlipFill_NoImage()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipFillNoImage");
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);
            shape.Fill.Style = eFillStyle.BlipFill;

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
        }
        [TestMethod]
        public void BlipFill_Stretch()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipFillStretch");

            var shape = AddBlip(ws, 1, "Shape1", false, 0, 0);
            shape.Fill.BlipFill.Stretch = true;
            shape.Fill.BlipFill.StretchOffset.TopOffset = 20;
            shape.Fill.BlipFill.StretchOffset.BottomOffset = 10;
            shape.Fill.BlipFill.StretchOffset.LeftOffset = -5;
            shape.Fill.BlipFill.StretchOffset.RightOffset = 15;

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
            Assert.AreEqual(true, shape.Fill.BlipFill.Stretch);
            Assert.AreEqual(20, shape.Fill.BlipFill.StretchOffset.TopOffset);
            Assert.AreEqual(10, shape.Fill.BlipFill.StretchOffset.BottomOffset);
            Assert.AreEqual(-5, shape.Fill.BlipFill.StretchOffset.LeftOffset);
            Assert.AreEqual(15, shape.Fill.BlipFill.StretchOffset.RightOffset);
        }
        [TestMethod]
        public void BlipFill_SourceRectangle()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipFillSourceRectangle");

            var shape = AddBlip(ws, 1, "Shape1", false, 0, 0);
            shape.Fill.BlipFill.Stretch = false;
            shape.Fill.BlipFill.SourceRectangle.TopOffset = 20;
            shape.Fill.BlipFill.SourceRectangle.BottomOffset = 10;
            shape.Fill.BlipFill.SourceRectangle.LeftOffset = -5;
            shape.Fill.BlipFill.SourceRectangle.RightOffset = 15;

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
            Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
            Assert.AreEqual(20, shape.Fill.BlipFill.SourceRectangle.TopOffset);
            Assert.AreEqual(10, shape.Fill.BlipFill.SourceRectangle.BottomOffset);
            Assert.AreEqual(-5, shape.Fill.BlipFill.SourceRectangle.LeftOffset);
            Assert.AreEqual(15, shape.Fill.BlipFill.SourceRectangle.RightOffset);
        }
        [TestMethod]
        public void BlipFill_Tile()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipFillTile");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.BlipFill;
            shape.Fill.BlipFill.Image.SetImage(Resources.CodeTif, ePictureType.Tif);

            shape.Fill.BlipFill.Stretch = false;
            shape.Fill.BlipFill.Tile.Alignment = eRectangleAlignment.Center;
            shape.Fill.BlipFill.Tile.FlipMode = eTileFlipMode.XY;
            shape.Fill.BlipFill.Tile.HorizontalRatio = 95;
            shape.Fill.BlipFill.Tile.VerticalRatio = 97;
            shape.Fill.BlipFill.Tile.HorizontalOffset = 2;
            shape.Fill.BlipFill.Tile.VerticalOffset = 1;

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
            Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
            Assert.AreEqual(eRectangleAlignment.Center, shape.Fill.BlipFill.Tile.Alignment);
            Assert.AreEqual(eTileFlipMode.XY, shape.Fill.BlipFill.Tile.FlipMode);
            Assert.AreEqual(95, shape.Fill.BlipFill.Tile.HorizontalRatio);
            Assert.AreEqual(97, shape.Fill.BlipFill.Tile.VerticalRatio);
            Assert.AreEqual(2, shape.Fill.BlipFill.Tile.HorizontalOffset);
            Assert.AreEqual(1, shape.Fill.BlipFill.Tile.VerticalOffset);
        }
        [TestMethod]
        public void BlipFill_PieChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BlipFillPieChart");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddPieChart("PieChart1", ePieChartType.Pie);
            var serie = chart.Series.Add("D2:D6", "A2:A6");
            chart.Fill.Style = eFillStyle.BlipFill;
            chart.Fill.BlipFill.Image.SetImage(Resources.CodeTif, ePictureType.Tif);


            var pt = chart.Series[0].DataPoints.Add(0);
            pt.Fill.Style = eFillStyle.BlipFill;
            pt.Fill.BlipFill.Image.SetImage(Resources.CodeTif, ePictureType.Tif);

            chart.SetPosition(1, 0, 5, 0);
        }
        [TestMethod]
        public void BlipFill_OverwriteImage()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlipOverwriteImage");

            var shape = AddBlip(ws, 1, "Shape1", false, 0, 0);

            //Act
            shape.Fill.Style = eFillStyle.BlipFill;
            shape.Fill.BlipFill.Image.SetImage(Resources.CodeTif, ePictureType.Tif);

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
        }

        private static ExcelShape AddBlip(ExcelWorksheet ws, int row, string shapeName, bool stretch, double offset, double sourceRect = 0)
        {
            var shape = ws.Drawings.AddShape(shapeName, eShapeStyle.RoundRect);
            shape.SetPosition(row, 0, 5, 0);

            //Act
            shape.Fill.Style = eFillStyle.BlipFill;
            shape.Fill.BlipFill.Image.SetImage(Resources.Test1JpgByteArray, ePictureType.Jpg);
            shape.Fill.BlipFill.Stretch = stretch;
            if (stretch)
            {
                shape.Fill.BlipFill.StretchOffset.TopOffset = offset;
                shape.Fill.BlipFill.StretchOffset.BottomOffset = offset;
                shape.Fill.BlipFill.StretchOffset.LeftOffset = offset;
                shape.Fill.BlipFill.StretchOffset.RightOffset = offset;
            }
            shape.Fill.BlipFill.SourceRectangle.TopOffset = sourceRect;
            shape.Fill.BlipFill.SourceRectangle.BottomOffset = sourceRect;
            shape.Fill.BlipFill.SourceRectangle.LeftOffset = sourceRect;
            shape.Fill.BlipFill.SourceRectangle.RightOffset = sourceRect;
            return shape;
        }
        #endregion

        [TestMethod]
        public void FillingOnlyBackgroundColorShouldApply()
        {
            var colorToSet = Color.Purple;

            using(var p = OpenPackage("unsettingColors.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("ColorGone");

                for (int i = 1; i < 20; i++)
                {
                    sheet.Cells[i, 14].Value = i;
                    sheet.Cells[i, 14].Style.Fill.PatternType = ExcelFillStyle.DarkGray;
                    sheet.Cells[i, 14].Style.Fill.PatternColor.SetColor(colorToSet);
                    sheet.Cells[i, 14].Style.Fill.BackgroundColor.SetColor(colorToSet);
                }

                SaveAndCleanup(p);
            }

            using (var p = OpenPackage("unsettingColors.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];

                Assert.AreEqual(sheet.Cells[1, 14].Style.Fill.BackgroundColor.Rgb, colorToSet.ToArgb().ToString("X"));
            }
        }
    }
}
