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
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.VBA;
using OfficeOpenXml.Drawing;

namespace EPPlusTest
{
    [TestClass]
    public class VBATests : TestBase
    {
        [Ignore]
        [TestMethod]
        public void ReadVBA()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\report.xlsm"));
            File.WriteAllText(@"c:\temp\vba\modules\dir.txt", package.Workbook.VbaProject.CodePage + "," + package.Workbook.VbaProject.Constants + "," + package.Workbook.VbaProject.Description + "," + package.Workbook.VbaProject.HelpContextID.ToString() + "," + package.Workbook.VbaProject.HelpFile1 + "," + package.Workbook.VbaProject.HelpFile2 + "," + package.Workbook.VbaProject.Lcid.ToString() + "," + package.Workbook.VbaProject.LcidInvoke.ToString() + "," + package.Workbook.VbaProject.LibFlags.ToString() + "," + package.Workbook.VbaProject.MajorVersion.ToString() + "," + package.Workbook.VbaProject.MinorVersion.ToString() + "," + package.Workbook.VbaProject.Name + "," + package.Workbook.VbaProject.ProjectID + "," + package.Workbook.VbaProject.SystemKind.ToString() + "," + package.Workbook.VbaProject.Protection.HostProtected.ToString() + "," + package.Workbook.VbaProject.Protection.UserProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VbeProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VisibilityState.ToString());
            foreach (var module in package.Workbook.VbaProject.Modules)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", module.Name), module.Code);
            }
            foreach (var r in package.Workbook.VbaProject.References)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", r.Name), r.Libid + " " + r.ReferenceRecordID.ToString());
            }

            List<X509Certificate2> ret = new List<X509Certificate2>();
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
            //package.Workbook.VbaProject.Protection.SetPassword("");
            package.SaveAs(new FileInfo(@"c:\temp\vbaSaved.xlsm"));
        }
        [Ignore]
        [TestMethod]
        public void Resign()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\vbaWrite.xlsm"));
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[11];
            package.SaveAs(new FileInfo(@"c:\temp\vbaWrite2.xlsm"));
        }
        [Ignore]
        [TestMethod]
        public void VbaError()
        {
            DirectoryInfo workingDir = new DirectoryInfo(@"C:\epplusExample\folder");
            if (!workingDir.Exists) workingDir.Create();
            FileInfo f = new FileInfo(workingDir.FullName + "//" + "temp.xlsx");
            if (f.Exists) f.Delete();
            ExcelPackage myPackage = new ExcelPackage(f);
            myPackage.Workbook.CreateVBAProject();
            ExcelWorksheet excelWorksheet = myPackage.Workbook.Worksheets.Add("Sheet1");
            ExcelWorksheet excelWorksheet2 = myPackage.Workbook.Worksheets.Add("Sheet2");
            ExcelWorksheet excelWorksheet3 = myPackage.Workbook.Worksheets.Add("Sheet3");
            FileInfo f2 = new FileInfo(workingDir.FullName + "//" + "newfile.xlsm");
            ExcelVBAModule excelVbaModule = myPackage.Workbook.VbaProject.Modules.AddModule("Module1");
            StringBuilder mybuilder = new StringBuilder(); mybuilder.AppendLine("Sub Jiminy()");
            mybuilder.AppendLine("Range(\"D6\").Select");
            mybuilder.AppendLine("ActiveCell.FormulaR1C1 = \"Jiminy\"");
            mybuilder.AppendLine("End Sub");
            excelVbaModule.Code = mybuilder.ToString();
            myPackage.SaveAs(f2);
            myPackage.Dispose();
        }
        [Ignore]
        [TestMethod]
        public void ReadVBAUnicodeWsName()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\VbaUnicodeWS.xlsm"));
            File.WriteAllText(@"c:\temp\vba\modules\dir.txt", package.Workbook.VbaProject.CodePage + "," + package.Workbook.VbaProject.Constants + "," + package.Workbook.VbaProject.Description + "," + package.Workbook.VbaProject.HelpContextID.ToString() + "," + package.Workbook.VbaProject.HelpFile1 + "," + package.Workbook.VbaProject.HelpFile2 + "," + package.Workbook.VbaProject.Lcid.ToString() + "," + package.Workbook.VbaProject.LcidInvoke.ToString() + "," + package.Workbook.VbaProject.LibFlags.ToString() + "," + package.Workbook.VbaProject.MajorVersion.ToString() + "," + package.Workbook.VbaProject.MinorVersion.ToString() + "," + package.Workbook.VbaProject.Name + "," + package.Workbook.VbaProject.ProjectID + "," + package.Workbook.VbaProject.SystemKind.ToString() + "," + package.Workbook.VbaProject.Protection.HostProtected.ToString() + "," + package.Workbook.VbaProject.Protection.UserProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VbeProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VisibilityState.ToString());
            foreach (var module in package.Workbook.VbaProject.Modules)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", module.Name), module.Code);
            }
            foreach (var r in package.Workbook.VbaProject.References)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", r.Name), r.Libid + " " + r.ReferenceRecordID.ToString());
            }

            List<X509Certificate2> ret = new List<X509Certificate2>();
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
            //package.Workbook.VbaProject.Protection.SetPassword("");
            package.SaveAs(new FileInfo(@"c:\temp\vbaSaved.xlsm"));
        }
        //Issue with chunk overwriting 4096 bytes
        [Ignore]
        [TestMethod]
        public void VbaBug()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\outfile.xlsm")))
            {
                Console.WriteLine(package.Workbook.CodeModule.Code.Length);
                package.Workbook.Worksheets[1].CodeModule.Code = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n\r\nEnd Sub";
                package.Workbook.Worksheets.Add("TestCopy", package.Workbook.Worksheets[1]);
                package.SaveAs(new FileInfo(@"c:\temp\bug\outfile2.xlsm"));
            }
        }
        [TestMethod]
        public void DecompressionChunkGreaterThan4k()
        {
            // This is a test for Issue 15026: VBA decompression encounters index out of range
            // on the decompression buffer.
            var workbookDir = Path.Combine(
#if Core
                AppContext.BaseDirectory
#else
                AppDomain.CurrentDomain.BaseDirectory
#endif
                , @"..\..\workbooks");
            var path = Path.Combine(workbookDir, "VBADecompressBug.xlsm");
            var f = new FileInfo(path);
            if (f.Exists)
            {
                using (var package = new ExcelPackage(f))
                {
                    // Reading the Workbook.CodeModule.Code will cause an IndexOutOfRange if the problem hasn't been fixed.
                    Assert.IsTrue(package.Workbook.CodeModule.Code.Length > 0);
                }
            }
        }
        [TestMethod, Ignore]
        public void ReadNewVBA()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\makro.xlsm")))
            {
                Console.WriteLine(package.Workbook.VbaProject.Modules[0].Name);
                
                package.SaveAs(new FileInfo(@"c:\temp\bug\makroepp.xlsm"));
            }
        }
        [TestMethod, Ignore]
        public void VBASigning()
        {
            using (var p = OpenPackage("vbaSign.xlsm", true))
            {
                p.Workbook.CreateVBAProject();
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Drawings.AddShape("Drawing1", eShapeStyle.Rect);

                //Now add some code to update the text of the shape...
                var sb = new StringBuilder();

                sb.AppendLine("Private Sub Workbook_Open()");
                sb.AppendLine("    [Test].Shapes(\"Drawing1\").TextEffect.Text = \"This text is set from VBA!\"");
                sb.AppendLine("End Sub");
                p.Workbook.CodeModule.Code = sb.ToString();

                //Optionally, Sign the code with your company certificate.
                X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                p.Workbook.VbaProject.Signature.Certificate = store.Certificates[4];

                SaveAndCleanup(p);
            }
        }
        [TestMethod, Ignore]
        public void VBASigningFromFile()
        {
            using (var p = OpenPackage("vbaSignFile.xlsm", true))
            {
                p.Workbook.CreateVBAProject();
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Drawings.AddShape("Drawing1", eShapeStyle.Rect);

                //Now add some code to update the text of the shape...
                var sb = new StringBuilder();

                sb.AppendLine("Private Sub Workbook_Open()");
                sb.AppendLine("    [Test].Shapes(\"Drawing1\").TextEffect.Text = \"This text is set from VBA!\"");
                sb.AppendLine("End Sub");
                p.Workbook.CodeModule.Code = sb.ToString();

                //Optionally, Sign the code with your company certificate.

                // Create a collection object and populate it using the PFX file
                X509Certificate2Collection collection = new X509Certificate2Collection();
                collection.Import("c:\\temp\\codecert.pfx", "EPPlus", X509KeyStorageFlags.PersistKeySet);

                p.Workbook.VbaProject.Signature.Certificate = collection[0];
                
                SaveAndCleanup(p);
            }
            }
    }
}