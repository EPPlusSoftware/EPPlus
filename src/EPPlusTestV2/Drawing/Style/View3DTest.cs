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
using System.IO;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class View3DTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Drawing3D.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            File.Copy(fileName, dirName + "\\Drawing3DRead.xlsx", true);
        }
        [TestMethod]
        public void Scene3dDefaultCamera()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Scene3DDefaultCamera");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.Camera.CameraType = ePresetCameraType.IsometricOffAxis1Left;

            //Assert
            Assert.AreEqual(ePresetCameraType.IsometricOffAxis1Left, shape.ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(eRigPresetType.ThreePt, shape.ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(eLightRigDirection.Top, shape.ThreeD.Scene.LightRig.Direction);
        }
        [TestMethod]
        public void Scene3dDefaultLightRigType()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Scene3DLightRigType");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.LightRig.RigType = eRigPresetType.Soft;

            //Assert
        }
        [TestMethod]
        public void Scene3dDefaultLightRigDir()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Scene3DLightRig");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.LightRig.Direction = eLightRigDirection.Top;

            //Assert
        }
        [TestMethod]
        public void Scene3dDefaultLightBackplane()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Scene3DBackplane");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.BackDropPlane.AnchorPoint.X = 3;
            //Assert
        }
        [TestMethod]
        public void View3dBevelBDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DBevelBDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.BottomBevel.BevelType = eBevelPresetType.Circle;
            //Assert
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Height);
            Assert.AreEqual(eBevelPresetType.Circle, shape.ThreeD.BottomBevel.BevelType);
        }
        [TestMethod]
        public void View3dBevelTDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DBevelTDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.TopBevel.Width=7;
            
            //Assert
            Assert.AreEqual(7, shape.ThreeD.TopBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.TopBevel.Height);
            Assert.AreEqual(eBevelPresetType.Circle, shape.ThreeD.TopBevel.BevelType);
        }
        [TestMethod]
        public void View3dContourColorDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DContourColorDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.ContourColor.SetSchemeColor(eSchemeColor.Accent6);
            //Assert
            Assert.AreEqual(eDrawingColorType.Scheme,shape.ThreeD.ContourColor.ColorType);
            Assert.AreEqual(eSchemeColor.Accent6, shape.ThreeD.ContourColor.SchemeColor.Color);
            Assert.AreEqual(1, shape.ThreeD.ContourWidth);
        }
        [TestMethod]
        public void View3dExtrusionColorDefault()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DExtrusionColorDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.ExtrusionColor.SetSchemeColor(eSchemeColor.Background1);
            //Assert

            Assert.AreEqual(eDrawingColorType.Scheme, shape.ThreeD.ExtrusionColor.ColorType);
            Assert.AreEqual(eSchemeColor.Background1, shape.ThreeD.ExtrusionColor.SchemeColor.Color);
            Assert.AreEqual(1, shape.ThreeD.ExtrusionHeight);
        }
        [TestMethod]
        public void View3dMaterialTypeDefault()
        {
            //Setup
            var expected = ePresetMaterialType.Plastic;
            var ws = _pck.Workbook.Worksheets.Add("View3DMaterialTypeDefault");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.MaterialType=expected;
            //Assert
            Assert.AreEqual(expected, shape.ThreeD.MaterialType);
        }
        [TestMethod]
        public void View3dNoScene()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DNoScene");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.MaterialType = ePresetMaterialType.Metal;
            shape.ThreeD.BottomBevel.BevelType = eBevelPresetType.ArtDeco;
            shape.ThreeD.BottomBevel.Width = 5;
            shape.ThreeD.TopBevel.BevelType = eBevelPresetType.Slope;
            shape.ThreeD.TopBevel.Height = 7;
            shape.ThreeD.ContourColor.SetSchemeColor(eSchemeColor.Accent4);
            shape.ThreeD.ContourWidth = 5;
            shape.ThreeD.ExtrusionColor.SetSystemColor(eSystemColor.Background);
            shape.ThreeD.ExtrusionHeight = 8;

            //Assert
            Assert.AreEqual(ePresetMaterialType.Metal, shape.ThreeD.MaterialType);

            Assert.AreEqual(eBevelPresetType.Slope, shape.ThreeD.TopBevel.BevelType);
            Assert.AreEqual(6, shape.ThreeD.TopBevel.Width);
            Assert.AreEqual(7, shape.ThreeD.TopBevel.Height);

            Assert.AreEqual(eBevelPresetType.ArtDeco, shape.ThreeD.BottomBevel.BevelType);
            Assert.AreEqual(5, shape.ThreeD.BottomBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Height);

            Assert.IsInstanceOfType(shape.ThreeD.ContourColor.SchemeColor, typeof(ExcelDrawingSchemeColor));
            Assert.AreEqual(eSchemeColor.Accent4, shape.ThreeD.ContourColor.SchemeColor.Color);
            Assert.AreEqual(5, shape.ThreeD.ContourWidth);
            Assert.IsInstanceOfType(shape.ThreeD.ExtrusionColor.SystemColor, typeof(ExcelDrawingSystemColor));
            Assert.AreEqual(eSystemColor.Background, shape.ThreeD.ExtrusionColor.SystemColor.Color);
            Assert.AreEqual(8, shape.ThreeD.ExtrusionHeight);
        }
        [TestMethod]
        public void View3dScene()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DScene");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.Camera.CameraType = ePresetCameraType.ObliqueTopLeft;
            shape.ThreeD.Scene.Camera.Zoom = 90;
            shape.ThreeD.Scene.Camera.Rotation.Latitude = 100;
            shape.ThreeD.Scene.Camera.Rotation.Longitude = 200;
            shape.ThreeD.Scene.Camera.Rotation.Revolution = 300;
            shape.ThreeD.Scene.LightRig.RigType=eRigPresetType.Sunset;
            shape.ThreeD.Scene.LightRig.Direction = eLightRigDirection.BottomRight;

            shape.ThreeD.MaterialType = ePresetMaterialType.LegacyWireframe;
            shape.ThreeD.BottomBevel.BevelType = eBevelPresetType.Divot;
            shape.ThreeD.BottomBevel.Height = 4;
            shape.ThreeD.TopBevel.BevelType = eBevelPresetType.Circle;
            shape.ThreeD.TopBevel.Width = 8;
            shape.ThreeD.ContourColor.SetHslColor(90, 50, 25);
            shape.ThreeD.ContourWidth = 4;
            shape.ThreeD.ExtrusionColor.SetPresetColor(ePresetColor.Azure);
            shape.ThreeD.ExtrusionHeight = 11;

            //Assert
            Assert.AreEqual(ePresetCameraType.ObliqueTopLeft, shape.ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(90, shape.ThreeD.Scene.Camera.Zoom);
            Assert.AreEqual(100, shape.ThreeD.Scene.Camera.Rotation.Latitude);
            Assert.AreEqual(200, shape.ThreeD.Scene.Camera.Rotation.Longitude);
            Assert.AreEqual(300, shape.ThreeD.Scene.Camera.Rotation.Revolution);
            Assert.AreEqual(eRigPresetType.Sunset, shape.ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(eLightRigDirection.BottomRight, shape.ThreeD.Scene.LightRig.Direction);
            

            Assert.AreEqual(ePresetMaterialType.LegacyWireframe, shape.ThreeD.MaterialType);

            Assert.AreEqual(eBevelPresetType.Circle, shape.ThreeD.TopBevel.BevelType);
            Assert.AreEqual(8, shape.ThreeD.TopBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.TopBevel.Height);

            Assert.AreEqual(eBevelPresetType.Divot, shape.ThreeD.BottomBevel.BevelType);
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Width);
            Assert.AreEqual(4, shape.ThreeD.BottomBevel.Height);

            Assert.AreEqual(eDrawingColorType.Hsl, shape.ThreeD.ContourColor.ColorType);
            Assert.AreEqual(90, shape.ThreeD.ContourColor.HslColor.Hue);
            Assert.AreEqual(4, shape.ThreeD.ContourWidth);
            Assert.AreEqual(eDrawingColorType.Preset, shape.ThreeD.ExtrusionColor.ColorType);
            Assert.AreEqual(ePresetColor.Azure, shape.ThreeD.ExtrusionColor.PresetColor.Color);
            Assert.AreEqual(11, shape.ThreeD.ExtrusionHeight);
        }
        [TestMethod]
        public void View3dNosp3d()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("View3DNosp3d");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.Camera.CameraType = ePresetCameraType.IsometricLeftDown;
            shape.ThreeD.Scene.Camera.Zoom = 80;
            shape.ThreeD.Scene.Camera.Rotation.Latitude = 5;
            shape.ThreeD.Scene.Camera.Rotation.Longitude = 10;
            shape.ThreeD.Scene.Camera.Rotation.Revolution = 20;
            shape.ThreeD.Scene.LightRig.RigType = eRigPresetType.LegacyNormal3;
            shape.ThreeD.Scene.LightRig.Direction = eLightRigDirection.TopLeft;
            shape.ThreeD.ShapeDepthZCoordinate = 3;

            //Assert
            Assert.AreEqual(ePresetCameraType.IsometricLeftDown, shape.ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(80, shape.ThreeD.Scene.Camera.Zoom);
            Assert.AreEqual(5, shape.ThreeD.Scene.Camera.Rotation.Latitude);
            Assert.AreEqual(10, shape.ThreeD.Scene.Camera.Rotation.Longitude);
            Assert.AreEqual(20, shape.ThreeD.Scene.Camera.Rotation.Revolution);
            Assert.AreEqual(eRigPresetType.LegacyNormal3, shape.ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(eLightRigDirection.TopLeft, shape.ThreeD.Scene.LightRig.Direction);
        }
        [TestMethod]
        public void Scene3dThreeD()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ThreeDTest");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.ThreeD.Scene.Camera.CameraType = ePresetCameraType.IsometricBottomDown;
            shape.ThreeD.Scene.LightRig.Direction = eLightRigDirection.Top;
            shape.ThreeD.Scene.LightRig.RigType = eRigPresetType.Sunset;
            shape.ThreeD.Scene.LightRig.Rotation.Revolution = 60;

            shape.ThreeD.MaterialType = ePresetMaterialType.DkEdge;
            shape.ThreeD.ContourWidth = 1;
            shape.ThreeD.ExtrusionHeight = 6;
            shape.ThreeD.ShapeDepthZCoordinate = 1;
            shape.ThreeD.TopBevel.Width = 0;
            shape.ThreeD.BottomBevel.BevelType = eBevelPresetType.RelaxedInset;
            shape.ThreeD.ExtrusionColor.SetSchemeColor(eSchemeColor.Background2);
            shape.ThreeD.ExtrusionColor.Transforms.AddLuminanceModulation(90);
            shape.ThreeD.ContourColor.SetSchemeColor(eSchemeColor.Accent1);
            shape.ThreeD.ContourColor.Transforms.AddLuminanceModulation(75);
            //Assert
        }
    }
}
