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
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class CommentsTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Comment.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

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
                a1.Comment.Fill.Style = OfficeOpenXml.Drawing.Vml.eVmlFillType.Gradient;
                a1.Comment.Fill.GradientSettings.SetGradientColors(
                    new OfficeOpenXml.Drawing.Vml.VmlGradiantColor(0, Color.Red), 
                    new OfficeOpenXml.Drawing.Vml.VmlGradiantColor(100, Color.Orange));
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
                Assert.IsNull(ws.Cells["A1"].Comment);
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
        [TestMethod]
        public void RangeShouldClearComment()
        {
            var ws = _pck.Workbook.Worksheets.Add("Sheet1");
            for (int i = 0; i < 5; i++)
            {
                ws.Cells[2, 2].Value = "hallo";
                ExcelComment comment = ws.Cells[2, 2].AddComment("hallo\r\nLine 2", "hallo");
                comment.Font.FontName = "Arial";
                comment.AutoFit = true;
                    
                ExcelRange cell = ws.Cells[2, 2];

                Assert.AreEqual("Arial", comment.Font.FontName);
                Assert.IsTrue(comment.AutoFit);
                Assert.AreEqual(1, ws.Comments.Count);
                Assert.IsNotNull(cell.Comment);

                cell.Clear();

                Assert.AreEqual(0, ws.Comments.Count);
                Assert.IsNull(cell.Comment);                                        
            }
        }
        [TestMethod]
        public void SettingRichTextShouldNotEffectComment()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelComment comment = ws.Cells[1, 1].AddComment("My Comment", "Me");
                Assert.IsNotNull(ws.Cells[1, 1].Comment);
                ws.Cells[1, 1].RichText.Add("RichText");
                Assert.IsNotNull(ws.Cells[1, 1].Comment);
            }
        }
        [TestMethod]
        public void CopyCommentInRange()
        {
            using (var p = new ExcelPackage())
            {
                // Get the comment object from the worksheet
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var comment1 = ws.Comments.Add(ws.Cells["B2"], "Test Comment");
                comment1.BackgroundColor = Color.FromArgb(0xdcf0ff);
                comment1.AutoFit = true;
                comment1.Font.FontName = "Tahoma";
                comment1.Font.Size = 9;
                comment1.Font.Bold = true; ;
                comment1.Font.Italic=true;
                comment1.Font.UnderLine = true;
                comment1.Font.Color = Color.FromArgb(0); 

                // Check that the comment in B2 has a custom style
                Assert.AreEqual("B2", comment1.Address);
                Assert.AreEqual("dcf0ff", comment1.BackgroundColor.Name);
                Assert.AreEqual(true, comment1.AutoFit);
                Assert.AreEqual("Tahoma", comment1.Font.FontName);
                Assert.AreEqual(9, comment1.Font.Size);
                Assert.AreEqual(true, comment1.Font.Bold);
                Assert.AreEqual(true, comment1.Font.Italic);
                Assert.AreEqual(true, comment1.Font.UnderLine);
                Assert.AreEqual("0", comment1.Font.Color.Name);

                // Copy the comment from B2 to A2 (also checking that this works when copying a range)
                ws.Cells["B1:B3"].Copy(ws.Cells["A1:A3"]);

                // Check the comment is copied with all properties intact
                var comment2 = ws.Comments[1];
                Assert.AreEqual("A2", comment2.Address);
                Assert.AreEqual(comment1.BackgroundColor.Name, comment2.BackgroundColor.Name);
                Assert.AreEqual(comment1.AutoFit, comment2.AutoFit);
                Assert.AreEqual(comment1.Font.FontName, comment2.Font.FontName);
                Assert.AreEqual(comment1.Font.Size, comment2.Font.Size);
                Assert.AreEqual(comment1.Font.Bold, comment2.Font.Bold);
                Assert.AreEqual(comment1.Font.Italic, comment2.Font.Italic);
                Assert.AreEqual(comment1.Font.UnderLine, comment2.Font.UnderLine);
                Assert.AreEqual(comment1.Font.Color.Name, comment2.Font.Color.Name);
            }
        }
        [TestMethod]
        public void TestDeleteCellsWithComment()
        {
            using (var p = new ExcelPackage())
            {
                // Add a sheet with comments
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Comments.Add(ws.Cells["B2"], "This is a comment.", "author");
                Assert.AreEqual(1, ws.Comments.Count);

                // Delete cells B1:B3 (including the comment in B2)
                ws.Cells["B1:B3"].Delete(eShiftTypeDelete.Left);

                // Check the comment is deleted
                Assert.AreEqual(0, ws.Comments.Count);
            }
        }
        [TestMethod]
        public void TestAddComments()
        {
            var ws = _pck.Workbook.Worksheets.Add("CommentSheet");
            ws.Cells[2, 2].Value = "hallo";
            ExcelComment comment = ws.Cells[2, 2].AddComment("hallo\r\nLine 2", "hallo");
            comment.Font.FontName = "Arial";
            comment.AutoFit = true;
        }
        [TestMethod]
        public void TestCreateSaveAndReadComment()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("CommentCreateSaveAndRead");
                var A1 = ws.Cells["A1"];
                A1.Value = "This cell has a ";
                var rt = A1.RichText.Add("comment");
                rt.Bold = true;
                rt.Italic = true;

                A1.AddComment("This is the ", "Myself");
                var comRt = A1.Comment.RichText.Add("comment");
                comRt.Bold = true;
                comRt.Italic = true;
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    var ws2 = p2.Workbook.Worksheets[0];
                    Assert.AreEqual("Myself", ws2.Cells["A1"].Comment.Author);
                    Assert.AreEqual("This is the comment", ws2.Cells["A1"].Comment.Text);
                    Assert.AreEqual("comment", ws2.Cells["A1"].Comment.RichText[1].Text);
                    Assert.IsTrue(ws2.Cells["A1"].Comment.RichText[1].Bold);
                    Assert.IsTrue(ws2.Cells["A1"].Comment.RichText[1].Italic);
                }
            }
        }
    }
}
