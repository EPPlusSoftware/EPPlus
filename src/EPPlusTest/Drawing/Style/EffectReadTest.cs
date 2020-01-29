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

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EffectReadTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DrawingEffectRead.xlsx");
        }
        [TestMethod]
        public void InnerShadowDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "InnerShadowDefault");

            //Assert default values
            Assert.AreEqual(0, shape.Effect.InnerShadow.BlurRadius);
            Assert.AreEqual(eDrawingColorType.Preset, shape.Effect.InnerShadow.Color.ColorType);
            Assert.AreEqual(ePresetColor.Black, shape.Effect.InnerShadow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.InnerShadow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.InnerShadow.Color.Transforms[0].Value);
            Assert.AreEqual(0, shape.Effect.InnerShadow.Direction);
            Assert.AreEqual(0, shape.Effect.InnerShadow.Distance);
        }
        [TestMethod]
        public void OuterShadowDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "OuterShadowDefault");

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
        public void ReflectionDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "ReflectionDefault");

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
        public void GlowDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "GlowDefault");

            //Assert default values
            Assert.AreEqual(0, shape.Effect.Glow.Radius);
            Assert.IsInstanceOfType(shape.Effect.Glow.Color.PresetColor, typeof(ExcelDrawingPresetColor));
            Assert.AreEqual(ePresetColor.Black, shape.Effect.Glow.Color.PresetColor.Color);

            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.Glow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.Glow.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void BlurDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "BlurDefault");
            
            //Assert default values
            Assert.AreEqual(0, shape.Effect.Blur.Radius);
            Assert.AreEqual(true, shape.Effect.Blur.GrowBounds);
        }
        [TestMethod]
        public void FillOverlayDefaultRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "FillOverlay");

            //Assert default values
            Assert.AreEqual(eBlendMode.Over,shape.Effect.FillOverlay.Blend);
        }

        [TestMethod]
        public void PresetShadowRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "PresetShadow");

            //Assert
            Assert.AreEqual(ePresetShadowType.FrontBottomShadow, shape.Effect.PresetShadow.Type);
            Assert.AreEqual(6,shape.Effect.PresetShadow.Direction);
            Assert.AreEqual(5, shape.Effect.PresetShadow.Distance);
            Assert.IsInstanceOfType(shape.Effect.PresetShadow.Color.PresetColor, typeof(ExcelDrawingPresetColor));
            Assert.AreEqual(ePresetColor.Black, shape.Effect.PresetShadow.Color.PresetColor.Color);
            Assert.AreEqual(eColorTransformType.Alpha, shape.Effect.PresetShadow.Color.Transforms[0].Type);
            Assert.AreEqual(50, shape.Effect.PresetShadow.Color.Transforms[0].Value);
        }
        [TestMethod]
        public void BlurRead()
        {
            //Setup
            var shape = TryGetShape(_pck, "Blur");

            //Assert
            Assert.AreEqual(50, shape.Effect.Blur.Radius);
            Assert.IsTrue(shape.Effect.Blur.GrowBounds);
            Assert.AreEqual("BlurMe", shape.Text);
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
