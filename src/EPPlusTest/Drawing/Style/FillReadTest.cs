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
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class FillReadTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingFillRead.xlsx");
        }
        #region SolidFill
        [TestMethod]
        public void ReadColorProperty()
        {
            //Setup
            var wsName = "SolidFill";
            var expected = Color.Blue;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
            Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        [TestMethod]
        public void ReadSolidFill_Color()
        {
            //Setup
            var wsName = "SolidFillFromSolidFill";
            var expected = Color.Green;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
            Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        [TestMethod]
        public void ReadSolidFill_ColorPreset()
        {
            //Setup
            var wsName = "SolidFillFromPresetClr";
            var expected = ePresetColor.Red;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.PresetColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.PresetColor.Color);
        }
        [TestMethod]
        public void ReadSolidFill_ColorScheme()
        {
            //Setup
            var wsName = "SolidFillFromSchemeClr";
            var expected = eSchemeColor.Accent6;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Act
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.SchemeColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SchemeColor.Color);
        }
        [TestMethod]
        public void ReadSolidFill_ColorPercentage()
        {
            //Setup
            var wsName = "SolidFillFromColorPrc";
            var expectedR = 51;
            var expectedG = 49;
            var expectedB = 50;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbPercentageColor);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.RgbPercentageColor.RedPercentage);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.RgbPercentageColor.GreenPercentage);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.RgbPercentageColor.BluePercentage);
        }
        [TestMethod]
        public void ReadSolidFill_ColorHsl()
        {
            //Setup
            var wsName = "SolidFillFromColorHcl";
            var expectedHue = 180;
            var expectedLum = 15;
            var expectedSat = 50;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.HslColor);
            Assert.AreEqual(expectedHue, shape.Fill.SolidFill.Color.HslColor.Hue);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.HslColor.Luminance);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.HslColor.Saturation);
        }
        [TestMethod]
        public void ReadSolidFill_ColorSystem()
        {
            //Setup
            var wsName = "SolidFillFromColorSystem";
            var expected = eSystemColor.Background;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.IsNotNull(shape.Fill.SolidFill.Color.SystemColor);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SystemColor.Color);
        }
        #endregion
        #region Transform
        [TestMethod]
        public void ReadTransparancy()
        {
            //Setup
            var wsName = "Transparancy";
            var expected = 45;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(expected, shape.Fill.Transparancy);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(100 - expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void ReadTransformAlpha()
        {
            //Setup
            var wsName = "Alpha";
            var expected = 45;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(100 - expected, shape.Fill.Transparancy);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void ReadTransformTint()
        {
            //Setup
            var wsName = "Tint";
            var expected = 30;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Tint, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void ReadTransformShade()
        {
            //Setup
            var wsName = "Shade";
            var expected = 95;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];


            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Shade, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void ReadTransformInverse_true()
        {
            //Setup
            var wsName = "Inverse_set";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Inv, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(1, shape.Fill.SolidFill.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void ReadTransformAlphaModulation()
        {
            //Setup
            var wsName = "AlphaModulation";
            var expected = 50;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(eColorTransformType.AlphaMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(20, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void ReadTransformAlphaOffset()
        {
            //Setup
            var wsName = "AlphaOffset";
            var expected = -10;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(eColorTransformType.AlphaOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(20, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void ReadTransformColorPercentage()
        {
            //Setup
            var wsName = "TransColorPerc";
            var expectedR = 30;
            var expectedG = 60;
            var expectedB = 20;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.Red, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.Green, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
            Assert.AreEqual(eColorTransformType.Blue, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void ReadTransformColorModulation()
        {
            //Setup
            var wsName = "TransColorMod";
            var expectedR = 3.33;
            var expectedG = 50;
            var expectedB = 25600;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.RedMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.GreenMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
            Assert.AreEqual(eColorTransformType.BlueMod, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void ReadTransformColorOffset()
        {
            //Setup
            var wsName = "TransColorOffset";
            var expectedR = 10;
            var expectedG = -20;
            var expectedB = 30;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.RedOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.GreenOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
            Assert.AreEqual(eColorTransformType.BlueOff, shape.Fill.SolidFill.Color.Transforms[2].Type);
            Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
        }
        [TestMethod]
        public void ReadTransformHslOffset()
        {
            //Setup
            var wsName = "TransHslOffset";
            var expectedLum = 10;
            var expectedSat = -20;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.LumOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.SatOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }
        [TestMethod]
        public void ReadTransformHslModulation()
        {
            //Setup
            var wsName = "TransHslModulation";
            var expectedLum = 50;
            var expectedSat = 200;
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];
            
            //Assert
            Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
            Assert.AreEqual(eColorTransformType.LumMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.SatMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
        }

        #endregion
        #region Gradient
        [TestMethod]
        public void ReadGradient()
        {
            //Setup
            var wsName = "Gradient";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);
            Assert.AreEqual(true, shape.Fill.GradientFill.RotateWithShape);
            Assert.AreEqual(eTileFlipMode.XY, shape.Fill.GradientFill.TileFlip);
        }
        [TestMethod]
        public void ReadGradientCircularPath()
        {
            //Setup
            var wsName = "GradientCircular";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);
            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.AreEqual(eShadePath.Circle, shape.Fill.GradientFill.ShadePath);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
        }
        [TestMethod]
        public void ReadGradientRectPath()
        {
            //Setup
            var wsName = "GradientRect";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        public void ReadGradientShapePath()
        {
            //Setup
            var wsName = "GradientShape";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);

            Assert.AreEqual(Color.LightBlue.ToArgb(), shape.Fill.GradientFill.Colors[0D].Color.RgbColor.Color.ToArgb());
            Assert.AreEqual(Color.Blue.ToArgb(), shape.Fill.GradientFill.Colors[40D].Color.RgbColor.Color.ToArgb());
            Assert.AreEqual(Color.DarkBlue.ToArgb(), shape.Fill.GradientFill.Colors[100D].Color.RgbColor.Color.ToArgb());
            Assert.IsNull(shape.Fill.GradientFill.Colors[41D]);

            Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
            Assert.AreEqual(eShadePath.Shape, shape.Fill.GradientFill.ShadePath);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
            Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
        }
        [TestMethod]
        public void ReadGradientAddMethods()
        {
            //Setup
            var wsName = "GradientAddMethods";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        #region Pattern
        [TestMethod]
        public void ReadPatternDefault()
        {
            //Setup
            var wsName = "PatternDefault";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.AreEqual(eFillPatternStyle.Pct5, shape.Fill.PatternFill.PatternType);
        }
        [TestMethod]
        public void ReadPatternCross()
        {
            //Setup
            var wsName = "PatternCross";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        #endregion
        #region Blip
        [TestMethod]
        public void ReadBlipFill_DefaultSettings()
        {
            //Setup
            var wsName = "BlipFill";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0]; 

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
            Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
        }
        [TestMethod]
        public void ReadBlipFill_NoImage()
        {
            //Setup
            var wsName = "BlipFillNoImage";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

            //Assert
            Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
            Assert.IsNull(shape.Fill.SolidFill);
            Assert.IsNull(shape.Fill.GradientFill);
            Assert.IsNull(shape.Fill.PatternFill);
        }
        [TestMethod]
        public void ReadBlipFill_Stretch()
        {
            //Setup
            var wsName = "BlipFillStretch";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        public void ReadBlipFill_SourceRectangle()
        {
            //Setup
            var wsName = "BlipFillSourceRectangle";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        public void ReadBlipFill_Tile()
        {
            //Setup
            var wsName = "BlipFillTile";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];

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
        public void ReadBlipFill_PieChart()
        {
            //Setup
            var wsName = "BlipFillPieChart";
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null)     Assert.Inconclusive($"{wsName} worksheet is missing");
            var chart = ws.Drawings[0].As.Chart.PieChart;

            Assert.AreEqual(eFillStyle.BlipFill, chart.Fill.Style);
            Assert.IsNull(chart.Fill.SolidFill);
            Assert.IsNull(chart.Fill.GradientFill);
            Assert.IsNull(chart.Fill.PatternFill);
            Assert.IsNotNull(chart.Fill.BlipFill.Image);
        }

        #endregion
    }
}
