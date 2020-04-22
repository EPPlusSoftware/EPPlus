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
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelChartAxisTest
    {
        private ExcelChartAxis axis;
        
        [TestInitialize]
        public void Initialize()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ></c:chartSpace>");
            var xmlNsm = new XmlNamespaceManager(new NameTable());
            xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
            xmlNsm.AddNamespace("a", ExcelPackage.schemaDrawings);
            var node = xmlDoc.CreateElement("axis");
            xmlDoc.DocumentElement.AppendChild(node);
            axis = new ExcelChartAxisStandard(null,xmlNsm, node, "c");
        }

        [TestMethod]
        public void CrossesAt_SetTo2_Is2()
        {
            axis.CrossesAt = 2;
            Assert.AreEqual(axis.CrossesAt, 2);
        }

        [TestMethod]
        public void CrossesAt_SetTo1EMinus6_Is1EMinus6()
        {
            axis.CrossesAt = 1.2e-6;
            Assert.AreEqual(axis.CrossesAt, 1.2e-6);
        }

        [TestMethod]
        public void MinValue_SetTo2_Is2()
        {
            axis.MinValue = 2;
            Assert.AreEqual(axis.MinValue, 2);
        }

        [TestMethod]
        public void MinValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MinValue = 1.2e-6;
            Assert.AreEqual(axis.MinValue, 1.2e-6);
        }

        [TestMethod]
        public void MaxValue_SetTo2_Is2()
        {
            axis.MaxValue = 2;
            Assert.AreEqual(axis.MaxValue, 2);
        }

        [TestMethod]
        public void MaxValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MaxValue = 1.2e-6;
            Assert.AreEqual(axis.MaxValue, 1.2e-6);
        }
        [TestMethod] 
        public void Gridlines_Set_IsNotNull()
        { 
            var major = axis.MajorGridlines;
            major.Width = 1;
            Assert.IsTrue(axis.ExistNode("c:majorGridlines")); 
  
            var minor = axis.MinorGridlines;
            minor.Width = 1;
            Assert.IsTrue(axis.ExistNode("c:minorGridlines")); 
        } 
  
        [TestMethod] 
        public void Gridlines_Remove_IsNull()
        { 
            var major = axis.MajorGridlines;
            major.Width = 1;
            var minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines(); 
  
            Assert.IsFalse(axis.ExistNode("c:majorGridlines")); 
            Assert.IsFalse(axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines;
            major.Width = 1;
            minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines(true, false); 
  
            Assert.IsFalse(axis.ExistNode("c:majorGridlines")); 
            Assert.IsTrue(axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines;
            major.Width = 1;
            minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines(false, true); 
  
            Assert.IsTrue(axis.ExistNode("c:majorGridlines"));
            Assert.IsFalse(axis.ExistNode("c:minorGridlines"));
        }
    }
}
