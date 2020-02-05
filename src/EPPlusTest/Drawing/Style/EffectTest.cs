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
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;
using System.IO;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EffectTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingEffect.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            File.Copy(fileName, dirName + "\\DrawingEffectRead.xlsx", true);
        }
        #region SetPreset Methods
        [TestMethod]
        public void SetPresetShadow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("PresetExcelShadow");

            //Act
            
            //Inner
            AddPresetShadowShape(ws, 1, 1, ePresetExcelShadowType.InnerTopLeft);
            AddPresetShadowShape(ws, 12, 1, ePresetExcelShadowType.InnerTop);
            AddPresetShadowShape(ws, 24, 1, ePresetExcelShadowType.InnerTopRight);
            AddPresetShadowShape(ws, 36, 1, ePresetExcelShadowType.InnerLeft);
            AddPresetShadowShape(ws, 48, 1, ePresetExcelShadowType.InnerCenter);
            AddPresetShadowShape(ws, 60, 1, ePresetExcelShadowType.InnerRight);
            AddPresetShadowShape(ws, 72, 1, ePresetExcelShadowType.InnerBottomLeft);
            AddPresetShadowShape(ws, 84, 1, ePresetExcelShadowType.InnerBottom);
            AddPresetShadowShape(ws, 96, 1, ePresetExcelShadowType.InnerBottomRight);

            //Outer
            AddPresetShadowShape(ws, 1, 13, ePresetExcelShadowType.OuterTopLeft);
            AddPresetShadowShape(ws, 12, 13, ePresetExcelShadowType.OuterTop);
            AddPresetShadowShape(ws, 24, 13, ePresetExcelShadowType.OuterTopRight);
            AddPresetShadowShape(ws, 36, 13, ePresetExcelShadowType.OuterLeft);
            AddPresetShadowShape(ws, 48, 13, ePresetExcelShadowType.OuterCenter);
            AddPresetShadowShape(ws, 60, 13, ePresetExcelShadowType.OuterRight);
            AddPresetShadowShape(ws, 72, 13, ePresetExcelShadowType.OuterBottomLeft);
            AddPresetShadowShape(ws, 84, 13, ePresetExcelShadowType.OuterBottom);
            AddPresetShadowShape(ws, 96, 13, ePresetExcelShadowType.OuterBottomRight);
            
            //Perspective
            AddPresetShadowShape(ws, 1, 26, ePresetExcelShadowType.PerspectiveUpperLeft);
            AddPresetShadowShape(ws, 12, 26, ePresetExcelShadowType.PerspectiveUpperRight);
            AddPresetShadowShape(ws, 24, 26, ePresetExcelShadowType.PerspectiveBelow);
            AddPresetShadowShape(ws, 36, 26, ePresetExcelShadowType.PerspectiveLowerLeft);
            AddPresetShadowShape(ws, 48, 26, ePresetExcelShadowType.PerspectiveLowerRight);

            //Perspective
            AddPresetShadowShape(ws, 1, 39, ePresetExcelShadowType.None);

            //Assert
        }
        [TestMethod]
        public void SetPresetReflection()
        {
            var ws = _pck.Workbook.Worksheets.Add("PresetExcelReflection");

            //Act
            
            AddPresetReflectionShape(ws, 1, 1, ePresetExcelReflectionType.TightTouching);
            AddPresetReflectionShape(ws, 20, 1, ePresetExcelReflectionType.Tight4Pt);
            AddPresetReflectionShape(ws, 40, 1, ePresetExcelReflectionType.Tight8Pt);

            AddPresetReflectionShape(ws, 1, 13, ePresetExcelReflectionType.HalfTouching);
            AddPresetReflectionShape(ws, 20, 13, ePresetExcelReflectionType.Half4Pt);
            AddPresetReflectionShape(ws, 40, 13, ePresetExcelReflectionType.Half8Pt);

            AddPresetReflectionShape(ws, 1, 26, ePresetExcelReflectionType.FullTouching);
            AddPresetReflectionShape(ws, 20, 26, ePresetExcelReflectionType.Full4Pt);
            AddPresetReflectionShape(ws, 40, 26, ePresetExcelReflectionType.Full8Pt);

            AddPresetReflectionShape(ws, 1, 39, ePresetExcelReflectionType.None);
        }
        [TestMethod]
        public void SetPresetSoftEdges()
        {
            var ws = _pck.Workbook.Worksheets.Add("PresetExcelSoftEdges");

            //Act

            AddPresetSoftEdgesShape(ws, 1, 1, ePresetExcelSoftEdgesType.None);
            AddPresetSoftEdgesShape(ws, 13, 1, ePresetExcelSoftEdgesType.SoftEdge1Pt);
            AddPresetSoftEdgesShape(ws, 26, 1, ePresetExcelSoftEdgesType.SoftEdge2_5Pt);
            AddPresetSoftEdgesShape(ws, 39, 1, ePresetExcelSoftEdgesType.SoftEdge5Pt);

            AddPresetSoftEdgesShape(ws, 1, 13, ePresetExcelSoftEdgesType.SoftEdge10Pt);
            AddPresetSoftEdgesShape(ws, 13, 13, ePresetExcelSoftEdgesType.SoftEdge25Pt);
            AddPresetSoftEdgesShape(ws, 26, 13, ePresetExcelSoftEdgesType.SoftEdge50Pt);
        }
        [TestMethod]
        public void SetPresetGlow()
        {
            var ws = _pck.Workbook.Worksheets.Add("PresetExcelGlow");

            //Act

            AddPresetGlowShape(ws, 1, 1, ePresetExcelGlowType.Accent1_5Pt);
            AddPresetGlowShape(ws, 13, 1, ePresetExcelGlowType.Accent1_8Pt);
            AddPresetGlowShape(ws, 26, 1, ePresetExcelGlowType.Accent1_11Pt);
            AddPresetGlowShape(ws, 39, 1, ePresetExcelGlowType.Accent1_18Pt);

            AddPresetGlowShape(ws, 1, 13, ePresetExcelGlowType.Accent2_5Pt);
            AddPresetGlowShape(ws, 13, 13, ePresetExcelGlowType.Accent2_8Pt);
            AddPresetGlowShape(ws, 26, 13, ePresetExcelGlowType.Accent2_11Pt);
            AddPresetGlowShape(ws, 39, 13, ePresetExcelGlowType.Accent2_18Pt);

            AddPresetGlowShape(ws, 1, 26, ePresetExcelGlowType.Accent3_5Pt);
            AddPresetGlowShape(ws, 13, 26, ePresetExcelGlowType.Accent3_8Pt);
            AddPresetGlowShape(ws, 26, 26, ePresetExcelGlowType.Accent3_11Pt);
            AddPresetGlowShape(ws, 39, 26, ePresetExcelGlowType.Accent3_18Pt);

            AddPresetGlowShape(ws, 1, 39, ePresetExcelGlowType.Accent4_5Pt);
            AddPresetGlowShape(ws, 13, 39, ePresetExcelGlowType.Accent4_8Pt);
            AddPresetGlowShape(ws, 26, 39, ePresetExcelGlowType.Accent4_11Pt);
            AddPresetGlowShape(ws, 39, 39, ePresetExcelGlowType.Accent4_18Pt);

            AddPresetGlowShape(ws, 1, 52, ePresetExcelGlowType.Accent5_5Pt);
            AddPresetGlowShape(ws, 13, 52, ePresetExcelGlowType.Accent5_8Pt);
            AddPresetGlowShape(ws, 26, 52, ePresetExcelGlowType.Accent5_11Pt);
            AddPresetGlowShape(ws, 39, 52, ePresetExcelGlowType.Accent5_18Pt);

            AddPresetGlowShape(ws, 1, 65, ePresetExcelGlowType.Accent6_5Pt);
            AddPresetGlowShape(ws, 13, 65, ePresetExcelGlowType.Accent6_8Pt);
            AddPresetGlowShape(ws, 26, 65, ePresetExcelGlowType.Accent6_11Pt);
            AddPresetGlowShape(ws, 39, 65, ePresetExcelGlowType.Accent6_18Pt);
        }
        #endregion
        [TestMethod]
        public void InnerShadowDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("InnerShadowDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.InnerShadow.BlurRadius=0;
            shape.Effect.InnerShadow.BlurRadius = null;

            //Assert default values
            Assert.AreEqual(0, shape.Effect.InnerShadow.BlurRadius);
            Assert.IsInstanceOfType(shape.Effect.InnerShadow.Color.PresetColor, typeof(ExcelDrawingPresetColor));
            Assert.AreEqual(ePresetColor.Black, shape.Effect.InnerShadow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.InnerShadow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.InnerShadow.Color.Transforms[0].Value);
            Assert.AreEqual(0, shape.Effect.InnerShadow.Direction);
            Assert.AreEqual(0, shape.Effect.InnerShadow.Distance);
        }
        [TestMethod]
        public void OuterShadowDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("OuterShadowDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.OuterShadow.Direction = 0;
            shape.Effect.OuterShadow.Direction = null;

            //Assert default values
            Assert.AreEqual(eRectangleAlignment.Bottom, shape.Effect.OuterShadow.Alignment);
            Assert.AreEqual(0, shape.Effect.OuterShadow.BlurRadius);
            Assert.AreEqual(eDrawingColorType.Preset, shape.Effect.OuterShadow.Color.ColorType);
            Assert.AreEqual(ePresetColor.Black, shape.Effect.OuterShadow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.OuterShadow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.OuterShadow.Color.Transforms[0].Value);
            Assert.AreEqual(0, shape.Effect.OuterShadow.Direction);
            Assert.AreEqual(0, shape.Effect.OuterShadow.Distance);
            Assert.AreEqual(100, shape.Effect.OuterShadow.HorizontalScalingFactor);
            Assert.AreEqual(0, shape.Effect.OuterShadow.HorizontalSkewAngle);
            Assert.AreEqual(true, shape.Effect.OuterShadow.RotateWithShape);
            Assert.AreEqual(100, shape.Effect.OuterShadow.VerticalScalingFactor);
            Assert.AreEqual(0, shape.Effect.OuterShadow.VerticalSkewAngle);
        }
        [TestMethod]
        public void ReflectionDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ReflectionDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.Reflection.VerticalSkewAngle=0;
            shape.Effect.Reflection.VerticalSkewAngle = null;

            //Assert default values
            Assert.AreEqual(eRectangleAlignment.Bottom, shape.Effect.Reflection.Alignment);
            Assert.AreEqual(0, shape.Effect.Reflection.Direction);
            Assert.AreEqual(0, shape.Effect.Reflection.Distance);
            Assert.AreEqual(0, shape.Effect.Reflection.EndOpacity);
            Assert.AreEqual(100, shape.Effect.Reflection.EndPosition);
            Assert.AreEqual(90, shape.Effect.Reflection.FadeDirection);
            Assert.AreEqual(100, shape.Effect.Reflection.HorizontalScalingFactor);
            Assert.AreEqual(0, shape.Effect.Reflection.HorizontalSkewAngle);
            Assert.AreEqual(true, shape.Effect.Reflection.RotateWithShape);
            Assert.AreEqual(100, shape.Effect.Reflection.StartOpacity);
            Assert.AreEqual(0, shape.Effect.Reflection.StartPosition);
            Assert.AreEqual(100, shape.Effect.Reflection.VerticalScalingFactor);
            Assert.AreEqual(0, shape.Effect.Reflection.VerticalSkewAngle);
        }
        [TestMethod]
        public void GlowDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("GlowDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.Glow.Radius = 0;
            shape.Effect.Glow.Radius = null;

            //Assert default values
            Assert.AreEqual(0, shape.Effect.Glow.Radius);
            Assert.AreEqual(eDrawingColorType.Preset, shape.Effect.Glow.Color.ColorType);
            Assert.AreEqual(ePresetColor.Black,shape.Effect.Glow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.Glow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.Glow.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void BlurDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BlurDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.Blur.Radius=0;
            shape.Effect.Blur.Radius = null;
            
            //Assert default values
            Assert.AreEqual(0, shape.Effect.Blur.Radius);
            Assert.AreEqual(true, shape.Effect.Blur.GrowBounds);
        }
        [TestMethod]
        public void FillOverlayDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("FillOverlay");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.FillOverlay.Blend=eBlendMode.Over;

            //Assert default values
            Assert.AreEqual(eBlendMode.Over,shape.Effect.FillOverlay.Blend);
        }

        [TestMethod]
        public void PresetShadow()
        {
            //Setup
            var expected = ePresetShadowType.FrontBottomShadow;
            var ws = _pck.Workbook.Worksheets.Add("PresetShadow");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.PresetShadow.Type = expected;
            shape.Effect.PresetShadow.Direction=6;
            shape.Effect.PresetShadow.Distance=5;            
            shape.Effect.PresetShadow.Color.SetPresetColor(ePresetColor.Black);
            shape.Effect.PresetShadow.Color.Transforms.AddAlpha(50);

            //Assert
            Assert.AreEqual(expected, shape.Effect.PresetShadow.Type);
            Assert.AreEqual(6,shape.Effect.PresetShadow.Direction);
            Assert.AreEqual(5, shape.Effect.PresetShadow.Distance);
            Assert.IsInstanceOfType(shape.Effect.PresetShadow.Color.PresetColor, typeof(ExcelDrawingPresetColor));
            Assert.AreEqual(ePresetColor.Black, shape.Effect.PresetShadow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.PresetShadow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.PresetShadow.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void Blur()
        {
            //Setup
            var expected = 50;
            var ws = _pck.Workbook.Worksheets.Add("Blur");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Effect.Blur.Radius = expected;
            shape.Effect.Blur.GrowBounds = true;
            shape.Text = "BlurMe";

            //Assert
            Assert.AreEqual(expected, shape.Effect.Blur.Radius);
            Assert.IsTrue(shape.Effect.Blur.GrowBounds);
        }

        #region Private help methods
        private static void AddPresetShadowShape(ExcelWorksheet ws, int row, int col, ePresetExcelShadowType preset)
        {
            var shape = AddShape(ws, row, col, preset.ToString());
            shape.Effect.SetPresetShadow(preset);
        }

        private static void AddPresetReflectionShape(ExcelWorksheet ws, int row, int col, ePresetExcelReflectionType preset)
        {
            var shape = AddShape(ws, row, col, preset.ToString());
            shape.Effect.SetPresetReflection(preset);
        }
        private static void AddPresetGlowShape(ExcelWorksheet ws, int row, int col, ePresetExcelGlowType preset)
        {
            var shape = AddShape(ws, row, col, preset.ToString());
            shape.Effect.SetPresetGlow(preset);
        }
        private static void AddPresetSoftEdgesShape(ExcelWorksheet ws, int row, int col, ePresetExcelSoftEdgesType preset)
        {
            var shape = AddShape(ws, row, col, preset.ToString());
            shape.Effect.SetPresetSoftEdges(preset);
        }
        private static ExcelShape AddShape(ExcelWorksheet ws, int row, int col, string name)
        {
            var shape = ws.Drawings.AddShape(name, eShapeStyle.RoundRect);
            shape.Text = name;
            shape.TextAlignment = eTextAlignment.Center;
            shape.TextAnchoring = eTextAnchoringType.Center;
            shape.Font.Color = Color.Black;
            shape.SetPosition(row, 0, col, 0);
            return shape;
        }
        #endregion
    }
}
