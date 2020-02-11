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
using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class CommentsTest
    {
        [TestMethod]
        public void VisibilityComments()
        {
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Comment");
                var a1 = ws.Cells["A1"];
                a1.Value = "Justin Dearing";
                a1.AddComment("I am A1s comment", "JD");
                Assert.IsFalse(a1.Comment.Visible); // Comments are by default invisible 
                a1.Comment.Visible = true;
                a1.Comment.Visible = false;
                Assert.IsNotNull(a1.Comment);
                //check style attribute
                var stylesDict = new System.Collections.Generic.Dictionary<string, string>();
                string[] styles = a1.Comment.Style
                    .Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                foreach(var s in styles)
                {
                    string[] split = s.Split(':');
                    if (split.Length == 2)
                    {
                        var k = (split[0] ?? "").Trim().ToLower();
                        var v = (split[1] ?? "").Trim().ToLower();
                        stylesDict[k] = v;
                    }
                }
                Assert.IsTrue(stylesDict.ContainsKey("visibility"));
                Assert.AreEqual("hidden", stylesDict["visibility"]);
                Assert.IsFalse(a1.Comment.Visible);
                    
                pkg.Save();
            }
        }
        [TestMethod]
        public void CommentInsertColumn()
        {
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("CommentInsert");
                ws.Cells["A1"].AddComment("na", "test");
                Assert.AreEqual(1, ws.Comments.Count);

                ws.InsertColumn(1, 1);

                Assert.AreEqual("B1", ws.Cells["B1"].Comment.Address);
                //Throws a null reference exception
                ws.Comments.Remove(ws.Cells["B1"].Comment);

                //Throws an exception "Comment does not exist"
                ws.DeleteColumn(2);
                Assert.AreEqual(0, ws.Comments.Count);
            }
        }
        [TestMethod]
        public void CommentDeleteColumn()
        {
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("CommentInsert");
                ws.Cells["B1"].AddComment("na", "test");
                Assert.AreEqual(1, ws.Comments.Count);

                ws.DeleteColumn(1, 1);

                Assert.AreEqual("A1", ws.Cells["A1"].Comment.Address);
                //Throws a null reference exception
                ws.Comments.Remove(ws.Cells["A1"].Comment);

                //Throws an exception "Comment does not exist"
                ws.DeleteColumn(1);
                Assert.AreEqual(0, ws.Comments.Count);
            }
        }
        [TestMethod]
        public void CommentInsertRow()
        {
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("CommentInsert");
                ws.Cells["A1"].AddComment("na", "test");
                Assert.AreEqual(1, ws.Comments.Count);

                ws.InsertRow(1, 1);

                Assert.AreEqual("A2", ws.Cells["A2"].Comment.Address);
                //Throws a null reference exception
                ws.Comments.Remove(ws.Cells["A2"].Comment);

                //Throws an exception "Comment does not exist"
                ws.DeleteRow(2);
                Assert.AreEqual(0, ws.Comments.Count);
            }
        }
        [TestMethod]
        public void CommentDeleteRow()
        {
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("CommentInsert");
                ws.Cells["A2"].AddComment("na", "test");
                Assert.AreEqual(1, ws.Comments.Count);

                ws.DeleteRow(1, 1);

                Assert.AreEqual("A1", ws.Cells["A1"].Comment.Address);
                //Throws a null reference exception
                ws.Comments.Remove(ws.Cells["A1"].Comment);

                //Throws an exception "Comment does not exist"
                ws.DeleteRow(1);
                Assert.AreEqual(0, ws.Comments.Count);
            }
        }

    }
}
