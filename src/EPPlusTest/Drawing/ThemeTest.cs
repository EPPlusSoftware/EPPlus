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
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class ThemeTest : TestBase
    {
        internal static ExcelPackage _pck;
        [TestMethod]
        public void PresetColors()
        {
            var l = new List<Color>();
            foreach (var name in Enum.GetNames(typeof(ePresetColor)))
            {
                var n = name;
                if (n.Contains("Grey")) n = n.Replace("Grey", "Gray");
                if (n.StartsWith("Dk")) n = n.Replace("Dk", "Dark");
                if (n.StartsWith("Dk")) n = n.Replace("Med", "Medium");
                if (n.StartsWith("Lt")) n = n.Replace("Lt", "Light");
                var c = Color.FromName(n);
                if (c == Color.Empty)
                {
                    Assert.Fail("Error Name");
                }
                l.Add(c);
            }
        }
        [TestMethod]
        public void SystemColor()
        {
            var l = new List<Color>();
            foreach (var name in Enum.GetNames(typeof(eSystemColor)))
            {
                var n = name;
                Color c = Color.Empty;
                foreach (var p in typeof(SystemColors).GetProperties(BindingFlags.Public | BindingFlags.Static))
                {
                    if (p.Name.Equals(n, StringComparison.CurrentCultureIgnoreCase))
                    {
                        c = (Color)p.GetValue(null, null);
                    }
                }

                if (c == Color.Empty)
                {
                    Console.WriteLine(n);
                }
                l.Add(c);
            }
        }
        [TestMethod]
        public void Read()
        {
            _pck = OpenTemplatePackage("Theme.xlsx");
            var theme = _pck.Workbook.ThemeManager;
            Assert.AreNotEqual(theme.CurrentTheme, null);
            Assert.AreEqual(theme.CurrentTheme.ColorScheme.Accent1.ColorType, eDrawingColorType.Rgb);
            Assert.AreEqual((uint)(theme.CurrentTheme.ColorScheme.Accent1.RgbColor.Color.ToArgb()), (uint)0xFF4472C4);
            Assert.AreEqual((uint)(theme.CurrentTheme.ColorScheme.Light1.SystemColor.LastColor.ToArgb()), (uint)0xFFFFFFFF);
            Assert.AreEqual(theme.CurrentTheme.ColorScheme.Light1.SystemColor.Color, eSystemColor.Window);

            Assert.AreEqual(theme.CurrentTheme.FontScheme.MajorFont.Count, 50);
            Assert.AreEqual(theme.CurrentTheme.FontScheme.MinorFont.Count, 50);

            theme.CurrentTheme.ColorScheme.Accent1.SetRgbPercentageColor(49, 50, 51);
            theme.CurrentTheme.ColorScheme.Accent2.SetHslColor(93, 50, 35);
            theme.CurrentTheme.ColorScheme.Accent3.SetPresetColor(ePresetColor.Azure);
            theme.CurrentTheme.ColorScheme.Accent4.SetPresetColor(ePresetColor.CornflowerBlue);
            theme.CurrentTheme.ColorScheme.Accent5.SetSystemColor(eSystemColor.DarkShadow3d);
            theme.CurrentTheme.ColorScheme.Accent6.SetRgbColor(Color.FromArgb(34, 34, 34));
            theme.CurrentTheme.ColorScheme.Accent1.Transforms.AddAlpha(50);

            var f1 = theme.CurrentTheme.FormatScheme.FillStyle[0];
            var f2 = theme.CurrentTheme.FormatScheme.FillStyle[1];
            var f3 = theme.CurrentTheme.FormatScheme.FillStyle[2];

            var b1 = theme.CurrentTheme.FormatScheme.BackgroundFillStyle[0];
            var b2 = theme.CurrentTheme.FormatScheme.BackgroundFillStyle[1];
            var b3 = theme.CurrentTheme.FormatScheme.BackgroundFillStyle[2];

            Assert.AreEqual(eFillStyle.GradientFill, f2.Style);
            foreach (var f in f2.GradientFill.Colors)
            {

            }
            Assert.AreEqual(eFillStyle.GradientFill, f3.Style);
            foreach (var f in f3.GradientFill.Colors)
            {

            }

            theme.CurrentTheme.FormatScheme.FillStyle[0].Style = eFillStyle.GradientFill;
            theme.CurrentTheme.FormatScheme.FillStyle[0].GradientFill.Colors.AddPreset(0, ePresetColor.LightSalmon);
            theme.CurrentTheme.FormatScheme.FillStyle[0].GradientFill.Colors.AddPreset(50, ePresetColor.Red);
            theme.CurrentTheme.FormatScheme.FillStyle[0].GradientFill.Colors.AddPreset(100, ePresetColor.DarkRed);

            Assert.AreEqual(theme.CurrentTheme.FormatScheme.BorderStyle[0].Cap, eLineCap.Flat);
            theme.CurrentTheme.FormatScheme.BorderStyle[0].Cap = eLineCap.Square;
            Assert.AreEqual(theme.CurrentTheme.FormatScheme.BorderStyle[0].Cap, eLineCap.Square);
            Assert.AreEqual(theme.CurrentTheme.FormatScheme.BorderStyle[0].Join, eLineJoin.Miter);
            Assert.AreEqual(theme.CurrentTheme.FormatScheme.BorderStyle[0].MiterJoinLimit, 800);
            theme.CurrentTheme.FormatScheme.BorderStyle[0].Join = eLineJoin.Bevel;
            theme.CurrentTheme.FormatScheme.BorderStyle[2].Alignment = ePenAlignment.Center;
            theme.CurrentTheme.FormatScheme.BorderStyle[1].TailEnd.Style = eEndStyle.Stealth;
            Assert.AreEqual(eEndStyle.Stealth, theme.CurrentTheme.FormatScheme.BorderStyle[1].TailEnd.Style);
            theme.CurrentTheme.FormatScheme.BorderStyle[1].HeadEnd.Style = eEndStyle.Arrow;
            Assert.AreEqual(eEndStyle.Arrow, theme.CurrentTheme.FormatScheme.BorderStyle[1].HeadEnd.Style);

            theme.CurrentTheme.FormatScheme.EffectStyle[0].Effect.Blur.Radius = 15;

            _pck.Workbook.DefaultThemeVersion = null;
            SaveWorkbook("Theme-saved.xlsx", _pck);
            _pck.Dispose();
        }
        [TestMethod]
        public void LoadThmx_ColorScheme()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.TestThemeThmx);
            _pck.Workbook.ThemeManager.CurrentTheme.ColorScheme.Light1.SetRgbColor(Color.DarkGray);
            _pck.Workbook.ThemeManager.CurrentTheme.ColorScheme.Dark1.SetRgbColor(Color.LightGray);
            _pck.Workbook.ThemeManager.CurrentTheme.ColorScheme.Dark2.SetRgbColor(Color.Red);
            _pck.Workbook.ThemeManager.CurrentTheme.ColorScheme.Light2.SetPresetColor(Color.Green);
            _pck.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetPresetColor(Color.Orange);
            var ws = _pck.Workbook.Worksheets.Add("ThemeLoaded");

            ws.Cells["A1:A12"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells["A1"].Value = "Background1";
            ws.Cells["A1"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Background1;
            ws.Cells["A2"].Value = "Text1";
            ws.Cells["A2"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Text1;
            ws.Cells["A3"].Value = "Background2";
            ws.Cells["A3"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Background2;
            ws.Cells["A4"].Value = "Text2";
            ws.Cells["A4"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Text2;
            ws.Cells["A5"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent1;
            ws.Cells["A6"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent2;
            ws.Cells["A7"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent3;
            ws.Cells["A8"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent4;
            ws.Cells["A9"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent5;
            ws.Cells["A10"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Accent6;
            ws.Cells["A11"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.Hyperlink;
            ws.Cells["A12"].Style.Fill.BackgroundColor.Theme = eThemeSchemeColor.FollowedHyperlink;

            var shape = ws.Drawings.AddShape("Oct", eShapeStyle.Octagon);
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Text1);

            /*** Assert ***/
            Assert.AreEqual(eThemeSchemeColor.Background1, ws.Cells["A1"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Text1, ws.Cells["A2"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Background2, ws.Cells["A3"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Text2, ws.Cells["A4"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent1, ws.Cells["A5"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent2, ws.Cells["A6"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent3, ws.Cells["A7"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent4, ws.Cells["A8"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent5, ws.Cells["A9"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Accent6, ws.Cells["A10"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.Hyperlink, ws.Cells["A11"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(eThemeSchemeColor.FollowedHyperlink, ws.Cells["A12"].Style.Fill.BackgroundColor.Theme);

            SaveWorkbook("ThemeLoaded.xlsx", _pck);
        }
        [TestMethod]
        public void LoadThmx_FontScheme()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.TestThemeThmx);
            Assert.IsNotNull(_pck.Workbook.ThemeManager.CurrentTheme);
            var currentTheme = _pck.Workbook.ThemeManager.CurrentTheme;

            /*** Assert ***/
            //Font scheme
            Assert.AreEqual(33, currentTheme.FontScheme.MajorFont.Count);
            Assert.AreEqual(33, currentTheme.FontScheme.MinorFont.Count);
            Assert.AreEqual("Century Gothic", currentTheme.FontScheme.MajorFont[0].Typeface);
            Assert.AreEqual("Century Gothic", currentTheme.FontScheme.MinorFont[0].Typeface);

            Assert.AreEqual("Sylfaen", currentTheme.FontScheme.MajorFont[32].Typeface);
            Assert.AreEqual("Sylfaen", currentTheme.FontScheme.MinorFont[32].Typeface);
        }
        [TestMethod]
        public void LoadThmx_BordersScheme()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.TestThemeThmx);
            Assert.IsNotNull(_pck.Workbook.ThemeManager.CurrentTheme);
            var currentTheme = _pck.Workbook.ThemeManager.CurrentTheme;

            /*** Assert ***/
            //Border scheme
            Assert.AreEqual(3, currentTheme.FormatScheme.BorderStyle.Count);

            Assert.AreEqual(eLineStyle.Solid, currentTheme.FormatScheme.BorderStyle[0].Style);
            Assert.AreEqual(ePenAlignment.Center, currentTheme.FormatScheme.BorderStyle[0].Alignment);
            Assert.AreEqual(eLineCap.Flat, currentTheme.FormatScheme.BorderStyle[0].Cap);
            Assert.AreEqual(eCompundLineStyle.Single, currentTheme.FormatScheme.BorderStyle[0].CompoundLineStyle);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].HeadEnd.Width);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].HeadEnd.Height);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].HeadEnd.Style);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].TailEnd.Width);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].TailEnd.Height);
            Assert.IsNull(currentTheme.FormatScheme.BorderStyle[0].TailEnd.Style);

            Assert.AreEqual(9525, currentTheme.FormatScheme.BorderStyle[0].Width);
            Assert.AreEqual(12700, currentTheme.FormatScheme.BorderStyle[1].Width);
            Assert.AreEqual(19050, currentTheme.FormatScheme.BorderStyle[2].Width);
        }
        [TestMethod]
        public void LoadThmx_Effect()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.TestThemeThmx);
            Assert.IsNotNull(_pck.Workbook.ThemeManager.CurrentTheme);
            var currentTheme = _pck.Workbook.ThemeManager.CurrentTheme;

            /*** Assert ***/
            /*
<a:effectStyleLst>
<a:effectStyle>
    <a:effectLst/>
</a:effectStyle>
<a:effectStyle>
    <a:effectLst/>
    <a:scene3d>
<a:camera prst="orthographicFront">
<a:rot rev="0" lon="0" lat="0"/>
</a:camera>
<a:lightRig dir="t" rig="threePt"/>
</a:scene3d>
<a:sp3d>
<a:bevelT w="25400" h="12700"/>
</a:sp3d>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw dir="5400000" rotWithShape="0" algn="ctr" dist="19050" blurRad="57150">
<a:srgbClr val="000000">
<a:alpha val="48000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot rev="0" lon="0" lat="0"/>
</a:camera>
<a:lightRig dir="t" rig="threePt"/>
</a:scene3d>
<a:sp3d>
<a:bevelT w="50800" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>             
             */

            //Effect scheme
            Assert.AreEqual(3, currentTheme.FormatScheme.EffectStyle.Count);

            Assert.AreEqual(ePresetCameraType.OrthographicFront, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.Camera.Rotation.Longitude);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.Camera.Rotation.Latitude);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.Camera.Rotation.Revolution);
            Assert.AreEqual(eLightRigDirection.Top, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.LightRig.Direction);
            Assert.AreEqual(eRigPresetType.ThreePt, currentTheme.FormatScheme.EffectStyle[1].ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(2, currentTheme.FormatScheme.EffectStyle[1].ThreeD.TopBevel.Width);
            Assert.AreEqual(1, currentTheme.FormatScheme.EffectStyle[1].ThreeD.TopBevel.Height);

            Assert.IsTrue(currentTheme.FormatScheme.EffectStyle[2].Effect.HasOuterShadow);
            Assert.AreEqual(eDrawingColorType.Rgb, currentTheme.FormatScheme.EffectStyle[2].Effect.OuterShadow.Color.ColorType);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[2].Effect.OuterShadow.Color.RgbColor.Color.ToArgb() & 0xFFFFFF);
            Assert.AreEqual(1, currentTheme.FormatScheme.EffectStyle[2].Effect.OuterShadow.Color.Transforms.Count);

            Assert.AreEqual(ePresetCameraType.OrthographicFront, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.Camera.CameraType);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.Camera.Rotation.Longitude);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.Camera.Rotation.Latitude);
            Assert.AreEqual(0, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.Camera.Rotation.Revolution);
            Assert.AreEqual(eLightRigDirection.Top, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.LightRig.Direction);
            Assert.AreEqual(eRigPresetType.ThreePt, currentTheme.FormatScheme.EffectStyle[2].ThreeD.Scene.LightRig.RigType);
            Assert.AreEqual(4, currentTheme.FormatScheme.EffectStyle[2].ThreeD.TopBevel.Width);
            Assert.AreEqual(2, currentTheme.FormatScheme.EffectStyle[2].ThreeD.TopBevel.Height);
        }

        [TestMethod]
        public void LoadThmx_FormatScheme_Fills()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.TestThemeThmx);
            Assert.IsNotNull(_pck.Workbook.ThemeManager.CurrentTheme);
            var currentTheme = _pck.Workbook.ThemeManager.CurrentTheme;

            /*** Assert ***/
            //Background Fill scheme
            Assert.AreEqual(3, currentTheme.FormatScheme.BackgroundFillStyle.Count);

            Assert.AreEqual(eFillStyle.SolidFill, currentTheme.FormatScheme.BackgroundFillStyle[0].Style);
            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.BackgroundFillStyle[0].SolidFill.Color.ColorType);
            Assert.AreEqual(eSchemeColor.Style, currentTheme.FormatScheme.BackgroundFillStyle[0].SolidFill.Color.SchemeColor.Color);

            Assert.AreEqual(eFillStyle.SolidFill, currentTheme.FormatScheme.BackgroundFillStyle[1].Style);
            Assert.AreEqual(eSchemeColor.Style, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.SchemeColor.Color);
            Assert.AreEqual(2, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.Transforms.Count);
            Assert.AreEqual(eColorTransformType.Tint, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(95, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(eColorTransformType.SatMod, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.Transforms[1].Type);
            Assert.AreEqual(170, currentTheme.FormatScheme.BackgroundFillStyle[1].SolidFill.Color.Transforms[1].Value);


            Assert.AreEqual(eFillStyle.GradientFill, currentTheme.FormatScheme.BackgroundFillStyle[2].Style);
            Assert.AreEqual(3, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors.Count);
            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[0].Color.ColorType);
            Assert.AreEqual(4, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[0].Color.Transforms.Count);

            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[1].Color.ColorType);
            Assert.AreEqual(50, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[50D].Position);
            Assert.AreEqual(4, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[1].Color.Transforms.Count);

            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[2].Color.ColorType);
            Assert.AreEqual(2, currentTheme.FormatScheme.BackgroundFillStyle[2].GradientFill.Colors[2].Color.Transforms.Count);

            //Fill styles
            Assert.AreEqual(3, currentTheme.FormatScheme.FillStyle.Count);

            Assert.AreEqual(eFillStyle.SolidFill, currentTheme.FormatScheme.FillStyle[0].Style);
            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.FillStyle[0].SolidFill.Color.ColorType);
            Assert.AreEqual(eSchemeColor.Style, currentTheme.FormatScheme.FillStyle[0].SolidFill.Color.SchemeColor.Color);

            Assert.AreEqual(eFillStyle.GradientFill, currentTheme.FormatScheme.FillStyle[1].Style);
            Assert.AreEqual(3, currentTheme.FormatScheme.FillStyle[1].GradientFill.Colors.Count);
            Assert.AreEqual(eDrawingColorType.Scheme, currentTheme.FormatScheme.FillStyle[1].GradientFill.Colors[0].Color.ColorType);
            Assert.AreEqual(4, currentTheme.FormatScheme.FillStyle[1].GradientFill.Colors[0].Color.Transforms.Count);
            Assert.AreEqual(3, currentTheme.FormatScheme.FillStyle[1].GradientFill.Colors[1].Color.Transforms.Count);
            Assert.AreEqual(3, currentTheme.FormatScheme.FillStyle[1].GradientFill.Colors[2].Color.Transforms.Count);
        }
        [TestMethod]
        public void ReadThmx()
        {
            _pck = OpenPackage("ThemeLoaded.xlsx");
            Assert.IsNotNull(_pck.Workbook.ThemeManager);
            Assert.IsNotNull(_pck.Workbook.ThemeManager.CurrentTheme);
        }

        #region Theme Savon
        [TestMethod]
        public void ValidateThemeSavonWithBlipFill()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.SavonThmx);

            var theme = _pck.Workbook.ThemeManager.CurrentTheme;
            Assert.AreEqual(eFillStyle.BlipFill, theme.FormatScheme.BackgroundFillStyle[2].Style);
            var ws=_pck.Workbook.Worksheets.Add("ThemeTest");
            LoadTestdata(ws);
            var chart = ws.Drawings.AddBarChart("ThisChart", eBarChartType.BarClustered3D);

            chart.Series.Add("D2:D8", "A2:A8");
            chart.Series.Add("B2:B8", "A2:A8");
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Bar3dChartStyle1);
            SaveWorkbook("ThemeSavonBlipFill.xlsx", _pck);
        }
        [TestMethod]
        public void ValidateThemeWoodTypeWithBlipFill()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.Load(Resources.WoodTypeThmx);

            var theme = _pck.Workbook.ThemeManager.CurrentTheme;
            Assert.AreEqual(eFillStyle.BlipFill, theme.FormatScheme.BackgroundFillStyle[2].Style);
            var ws=_pck.Workbook.Worksheets.Add("ThemeTest");
            LoadTestdata(ws); 
            SaveWorkbook("ThemeWoodTypeBlipFill.xlsx", _pck);
        }
        [TestMethod]
        public void ChangeFontOnDefaultTheme()
        {
            _pck = new ExcelPackage();
            _pck.Workbook.ThemeManager.CreateDefaultTheme();

            var theme = _pck.Workbook.ThemeManager.CurrentTheme;
            Assert.AreEqual("Calibri Light", theme.FontScheme.MajorFont[0].Typeface);
            Assert.AreEqual("Calibri", theme.FontScheme.MinorFont[0].Typeface);
            theme.Name = "My custom theme";
            theme.FontScheme.MajorFont.SetLatinFont("Arial");
            theme.FontScheme.MinorFont.SetLatinFont("Arial");

            Assert.AreEqual("Arial", theme.FontScheme.MajorFont[0].Typeface);
            Assert.AreEqual("Arial", theme.FontScheme.MinorFont[0].Typeface);

            //Set normal font to arial
            _pck.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";

            _pck.Workbook.Worksheets.Add("Sheet1");
            SaveWorkbook("DefaultTheme.xlsx", _pck);
        }

        #endregion
    }
}   