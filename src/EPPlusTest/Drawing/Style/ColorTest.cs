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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;

namespace EPPlusTest.Drawing.Style
{
    [TestClass]
    public class ColorTest 
    {
        [TestMethod]
        public void VerifyPresetColorEnumCastFromColor()
        {
            var t = typeof(System.Drawing.Color);

            foreach(var pi in t.GetProperties(BindingFlags.Static | BindingFlags.Public))
            {
                if (pi.Name.Equals("transparant", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (!Enum.TryParse(pi.Name, out ePresetColor v))
                    {
                        Assert.Fail($"Convert to ePresetColorFailed for {pi.Name}");
                    }
                }
            }
        }
        [TestMethod]
        public void VerifyAlphaPartWhenSetColor()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("DrawingAlphaSetColor");

            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);
            shape.SetPosition(1, 0, 5, 0);

            //Act
            shape.Fill.Style=eFillStyle.SolidFill;
            shape.Fill.SolidFill.Color.SetRgbColor(Color.FromArgb(127,255,0,0), true);

            //Assert
            Assert.AreEqual(50, shape.Fill.SolidFill.Color.Transforms[0].Value);
            Assert.AreEqual(0xFFFF0000, (uint)shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }
        private static string TranslateFromColor(Color c)
        {
            if (c.IsEmpty || c.GetType().GetProperty(c.Name, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static) == null)
            {
                throw (new ArgumentException("A preset color can not be set to empty or be an unnamed color"));
            }
            var s = c.Name.ToString();
            return s.Substring(0, 1).ToLower() + s.Substring(1);
        }
        [TestMethod]
        public void HslToRgb()
        {
            var rgb = ExcelDrawingHslColor.GetRgb(300, 1, 1);
            Assert.AreEqual(0xFFFFFFFF, (uint)rgb.ToArgb());

            rgb = ExcelDrawingHslColor.GetRgb(300, 1, 0);
            Assert.AreEqual(0xFF000000, (uint)rgb.ToArgb());

            rgb = ExcelDrawingHslColor.GetRgb(0, 1, .5);
            Assert.AreEqual(0xFFFF0000, (uint)rgb.ToArgb());

            //Lime
            rgb = ExcelDrawingHslColor.GetRgb(120, 1, .5);
            Assert.AreEqual(0xFF00FF00, (uint)rgb.ToArgb());

            //Blue
            rgb = ExcelDrawingHslColor.GetRgb(240, 1, .5);
            Assert.AreEqual(0xFF0000FF, (uint)rgb.ToArgb());

            //Yellow
            rgb = ExcelDrawingHslColor.GetRgb(60, 1, .5);
            Assert.AreEqual(0xFFFFFF00, (uint)rgb.ToArgb());

            //Cyan
            rgb = ExcelDrawingHslColor.GetRgb(180, 1, .5);
            Assert.AreEqual(0xFF00FFFF, (uint)rgb.ToArgb());

            //Magenta
            rgb = ExcelDrawingHslColor.GetRgb(300, 1, .5);
            Assert.AreEqual(0xFFFF00FF, (uint)rgb.ToArgb());

            //Silver
            rgb = ExcelDrawingHslColor.GetRgb(0, 0, .75);
            Assert.AreEqual(0xFFBFBFBF, (uint)rgb.ToArgb());

            //Gray
            rgb = ExcelDrawingHslColor.GetRgb(0, 0, .50);
            Assert.AreEqual(0xFF808080, (uint)rgb.ToArgb());

            //Maroon 
            rgb = ExcelDrawingHslColor.GetRgb(0, 1, .25);
            Assert.AreEqual(0xFF800000, (uint)rgb.ToArgb());

            //Olive 
            rgb = ExcelDrawingHslColor.GetRgb(0, 1, .25);
            Assert.AreEqual(0xFF800000, (uint)rgb.ToArgb());

            //Green
            rgb = ExcelDrawingHslColor.GetRgb(120, 1, .25);
            Assert.AreEqual(0xFF008000, (uint)rgb.ToArgb());

            //Purple
            rgb = ExcelDrawingHslColor.GetRgb(300, 1, .25);
            Assert.AreEqual(0xFF800080, (uint)rgb.ToArgb());

            //Teal
            rgb = ExcelDrawingHslColor.GetRgb(180, 1, .25);
            Assert.AreEqual(0xFF008080, (uint)rgb.ToArgb());

            //43, 58%, 73%
            rgb = ExcelDrawingHslColor.GetRgb(43, .58, .73);
            Assert.AreEqual(0xFFE2CB92, (uint)rgb.ToArgb());

            //359, 79%, 21%
            rgb = ExcelDrawingHslColor.GetRgb(359, .79, .21);
            Assert.AreEqual(0xFF600B0D, (uint)rgb.ToArgb());            
        }
    }
}
