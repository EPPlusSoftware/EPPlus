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
    public class View3DReadTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Drawing3DRead.xlsx");
        }
        [TestMethod]
        public void Scene3dDefaultCameraRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "Scene3DDefaultCamera");

            //Assert
            Assert.AreEqual(ePresetCameraType.IsometricOffAxis1Left, shape.ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(eRigPresetType.ThreePt, shape.ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(eLightRigDirection.Top, shape.ThreeD.Scene.LightRig.Direction);
        }

        [TestMethod]
        public void Scene3dDefaultLightRigTypeRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "Scene3DLightRigType");

            //Assert
            Assert.AreEqual(shape.ThreeD.Scene.LightRig.RigType, eRigPresetType.Soft);
        }
        [TestMethod]
        public void Scene3dDefaultLightRigDirRead()
        {
            ExcelShape shape = TryGetShape(_pck, "Scene3DLightRig");

            //Assert
            Assert.AreEqual(shape.ThreeD.Scene.LightRig.Direction, eLightRigDirection.Top);
        }
        [TestMethod]
        public void Scene3dDefaultLightBackplanReade()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "Scene3DBackplane");

            //Assert
            Assert.AreEqual(3, shape.ThreeD.Scene.BackDropPlane.AnchorPoint.X);
        }
        [TestMethod]
        public void View3dBevelBDefaultRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DBevelBDefault");

            //Assert
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Height);
            Assert.AreEqual(eBevelPresetType.Circle, shape.ThreeD.BottomBevel.BevelType);
        }
        [TestMethod]
        public void View3dBevelTDefaultRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DBevelTDefault");
           
            //Assert
            Assert.AreEqual(7, shape.ThreeD.TopBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.TopBevel.Height);
            Assert.AreEqual(eBevelPresetType.Circle, shape.ThreeD.TopBevel.BevelType);
        }
        [TestMethod]
        public void View3dContourColorDefaultRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DContourColorDefault");

            //Assert
            Assert.AreEqual(eDrawingColorType.Scheme, shape.ThreeD.ContourColor.ColorType);
            Assert.AreEqual(eSchemeColor.Accent6, shape.ThreeD.ContourColor.SchemeColor.Color);
            Assert.AreEqual(1, shape.ThreeD.ContourWidth);
        }
        [TestMethod]
        public void View3dExtrusionColorDefaultRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DExtrusionColorDefault");

            //Assert
            Assert.AreEqual(eDrawingColorType.Scheme, shape.ThreeD.ExtrusionColor.ColorType);
            Assert.AreEqual(eSchemeColor.Background1, shape.ThreeD.ExtrusionColor.SchemeColor.Color);
            Assert.AreEqual(1, shape.ThreeD.ExtrusionHeight);
        }
        [TestMethod]
        public void View3dMaterialTypeDefaultRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DMaterialTypeDefault");

            //Assert
            Assert.AreEqual(ePresetMaterialType.Plastic, shape.ThreeD.MaterialType);
        }
        [TestMethod]
        public void View3dNoSceneRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DNoScene");

            //Assert
            Assert.AreEqual(ePresetMaterialType.Metal, shape.ThreeD.MaterialType);

            Assert.AreEqual(eBevelPresetType.Slope, shape.ThreeD.TopBevel.BevelType);
            Assert.AreEqual(6, shape.ThreeD.TopBevel.Width);
            Assert.AreEqual(7, shape.ThreeD.TopBevel.Height);

            Assert.AreEqual(eBevelPresetType.ArtDeco, shape.ThreeD.BottomBevel.BevelType);
            Assert.AreEqual(5, shape.ThreeD.BottomBevel.Width);
            Assert.AreEqual(6, shape.ThreeD.BottomBevel.Height);

            Assert.AreEqual(eDrawingColorType.Scheme, shape.ThreeD.ContourColor.ColorType);
            Assert.AreEqual(eSchemeColor.Accent4, shape.ThreeD.ContourColor.SchemeColor.Color);
            Assert.AreEqual(5, shape.ThreeD.ContourWidth);
            Assert.AreEqual(eDrawingColorType.System, shape.ThreeD.ExtrusionColor.ColorType);
            Assert.AreEqual(eSystemColor.Background, shape.ThreeD.ExtrusionColor.SystemColor.Color);
            Assert.AreEqual(8, shape.ThreeD.ExtrusionHeight);
        }
        [TestMethod]
        public void View3dSceneRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DScene");

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
        public void View3dNosp3dRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "View3DNosp3d");

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
        public void Scene3dThreeDRead()
        {
            //Setup
            ExcelShape shape = TryGetShape(_pck, "ThreeDTest");

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
