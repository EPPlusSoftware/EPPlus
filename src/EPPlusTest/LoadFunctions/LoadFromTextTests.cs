/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
//using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromTextTests : TestBase
    {
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("test");
            _lines = new StringBuilder();
            _format = new ExcelTextFormat();
            _formatFixed = new ExcelTextFormatFixedWidth();
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        private StringBuilder _lines;
        private ExcelTextFormat _format;
        private ExcelTextFormatFixedWidth _formatFixed;

        private void AddLine(string s)
        {
            _lines.AppendLine(s);
        }

        [TestMethod]
        public void ShouldLoadCsvFormat()
        {
            AddLine("a,b,c");
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString());
            Assert.AreEqual("a", _worksheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldLoadCsvFormatWithDelimiter()
        {
            AddLine("a;b;c");
            AddLine("d;e;f");
            _format.Delimiter = ';';
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format);
            Assert.AreEqual("a", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("d", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldUseTypesFromFormat()
        {
            AddLine("a;2;5%");
            AddLine("d;3;8%");
            _format.Delimiter = ';';
            _format.DataTypes = new eDataTypes[] { eDataTypes.String, eDataTypes.Number, eDataTypes.Percent };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format);
            Assert.AreEqual("a", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(2d, _worksheet.Cells["B1"].Value);
            Assert.AreEqual(3d, _worksheet.Cells["B2"].Value);
            Assert.AreEqual(0.05, _worksheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void ShouldUseHeadersFromFirstRow()
        {
            AddLine("Height 1,Width");
            AddLine("1,2");
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format, TableStyles.None, true);
            Assert.AreEqual("Height 1", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(1d, _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldUseTextQualifier()
        {
            AddLine("'Look, a bird!',2");
            AddLine("'One apple, one orange',3");
            _format.TextQualifier = '\'';
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format);
            Assert.AreEqual("Look, a bird!", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("One apple, one orange", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldReturnRange()
        {
            AddLine("a,b,c");
            var r = _worksheet.Cells["A1"].LoadFromText(_lines.ToString());
            Assert.AreEqual("A1:C2", r.FirstAddress);
        }
        [TestMethod]
        public void VerifyOneLineWithTextQualifier()
        {
            AddLine("\"a\",\"\"\"\", \"\"\"\"");
            var r = _worksheet.Cells["A1"].LoadFromText(_lines.ToString(),new ExcelTextFormat { TextQualifier='\"' });
            Assert.AreEqual("a", _worksheet.Cells[1,1].Value);
            Assert.AreEqual("\"", _worksheet.Cells[1, 2].Value);
            Assert.AreEqual("\"", _worksheet.Cells[1, 3].Value);
        }
        [TestMethod]
        public void VerifyMultiLineWithTextQualifier()
        {
            AddLine("\"a\",b, \"c\"\"\"");
            AddLine("a,\"b\", \"c\"\"\r\n\"\"\"");
            AddLine("a,\"b\", \"c\"\"\"\"\"");
            AddLine("\"d\",e, \"\"");
            AddLine("\"\",, \"\"");

            var r = _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), new ExcelTextFormat { TextQualifier = '\"' });
            Assert.AreEqual("a", _worksheet.Cells[1, 1].Value);
            Assert.AreEqual("b", _worksheet.Cells[1, 2].Value);
            Assert.AreEqual("c\"", _worksheet.Cells[1, 3].Value);

            Assert.AreEqual("a", _worksheet.Cells[2, 1].Value);
            Assert.AreEqual("b", _worksheet.Cells[2, 2].Value);
            Assert.AreEqual("c\"\r\n\"", _worksheet.Cells[2, 3].Value);

            Assert.AreEqual("a", _worksheet.Cells[3, 1].Value);
            Assert.AreEqual("b", _worksheet.Cells[3, 2].Value);
            Assert.AreEqual("c\"\"", _worksheet.Cells[3, 3].Value);

            Assert.AreEqual("d", _worksheet.Cells[4, 1].Value);
            Assert.AreEqual("e", _worksheet.Cells[4, 2].Value);
            Assert.IsNull(_worksheet.Cells[4, 3].Value);

            Assert.IsNull(_worksheet.Cells[5, 1].Value);
            Assert.IsNull(_worksheet.Cells[5, 2].Value);
            Assert.IsNull(_worksheet.Cells[5, 3].Value);
        }

        [TestMethod]
        public void SelectedColumnsText()
        {
            AddLine("a,b,c");
            _format.UseColumns = new bool[] { true, false, true };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format);

            Assert.AreEqual("a", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("c", _worksheet.Cells["B1"].Value);
        }

        [TestMethod]
        public void UseRowText()
        {
            AddLine("a,b,c");
            AddLine("d,e,f");
            AddLine("g,h,i");
            _format.ShouldUseRow = row => {
                if (row.Contains("e"))
                {
                    return false;
                }
                return true;
            };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _format);

            Assert.AreEqual("a", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("i", _worksheet.Cells["C2"].Value);
        }

        [TestMethod]
        public void ShouldLoadFixedWidthText()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("Yes", _worksheet.Cells["L4"].Value);
        }

        [TestMethod]
        public void ShouldUseTypesFromFormatFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _formatFixed.ColumnFormat[0].DataType = eDataTypes.Number;
            _formatFixed.ColumnFormat[1].DataType = eDataTypes.Number;
            _formatFixed.ColumnFormat[2].DataType = eDataTypes.DateTime;
            _formatFixed.ColumnFormat[3].DataType = eDataTypes.Number;
            _formatFixed.ColumnFormat[4].DataType = eDataTypes.String;
            _formatFixed.ColumnFormat[5].DataType = eDataTypes.String;
            _formatFixed.ColumnFormat[6].DataType = eDataTypes.String;
            _formatFixed.ColumnFormat[7].DataType = eDataTypes.String;
            _formatFixed.ColumnFormat[8].DataType = eDataTypes.String;
            _formatFixed.ColumnFormat[9].DataType = eDataTypes.Number;
            _formatFixed.ColumnFormat[10].DataType = eDataTypes.Number;
            _formatFixed.ColumnFormat[11].DataType = eDataTypes.String;
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed, TableStyles.None, true);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(3930621977d, _worksheet.Cells["D2"].Value);
            Assert.AreEqual(5007.10, _worksheet.Cells["K2"].Value);
        }

        [TestMethod]
        public void ShouldUseHeadersFromFirstRowFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed, TableStyles.None, true);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(16524d, _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldReturnRangeFixedWidth()
        {
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            var r = _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);
            Assert.AreEqual("Alloc.", _worksheet.Cells["L1"].Value);
            Assert.AreEqual("A1:L1", r.FirstAddress);
        }

        [TestMethod]
        public void OnlyMiddleRowFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   "); //<-this is the row we check for
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.SkipLinesBeginning = 2;
            _formatFixed.SkipLinesEnd = 1;
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual(191675d, _worksheet.Cells["A1"].Value);
            Assert.AreEqual("OUNHQEX XUFQONY", _worksheet.Cells["E1"].Value);
        }

        [TestMethod]
        public void SelectedColumnsFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.SetColumns(FixedWidthReadType.Length, 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _formatFixed.ColumnFormat[0].UseColumn = false;
            _formatFixed.ColumnFormat[1].UseColumn = false;
            _formatFixed.ColumnFormat[2].UseColumn = true;
            _formatFixed.ColumnFormat[3].UseColumn = true;
            _formatFixed.ColumnFormat[4].UseColumn = false;
            _formatFixed.ColumnFormat[5].UseColumn = false;
            _formatFixed.ColumnFormat[6].UseColumn = false;
            _formatFixed.ColumnFormat[7].UseColumn = false;
            _formatFixed.ColumnFormat[8].UseColumn = false;
            _formatFixed.ColumnFormat[9].UseColumn = false;
            _formatFixed.ColumnFormat[10].UseColumn = true;
            _formatFixed.ColumnFormat[11].UseColumn = false;
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual(3930621977d, _worksheet.Cells["B2"].Value);
            Assert.AreEqual(43537.00, _worksheet.Cells["C3"].Value);
        }

        [TestMethod]
        public void ReadFromPositionFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.ReadType = FixedWidthReadType.Positions;
            int[] arr = { 0, 8, 13, 24, 37, 69, 75, 81, 91, 95, 115, 135 };
            _formatFixed.SetColumns(FixedWidthReadType.Positions, arr );
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(2368183100d, _worksheet.Cells["D3"].Value);
        }

        [TestMethod]
        public void ReadFileFromDiskAndSaveAsXLSXFixedWidth()
        {
            string myFile = File.ReadAllText("C:\\Users\\AdrianParnéus\\Documents\\Test\\FixedWidth.txt");
            FileInfo myFileInfo = new FileInfo("C:\\Users\\AdrianParnéus\\Documents\\Test\\FixedWidth.txt");
            using (var p = OpenTemplatePackage("Fixed.xlsx"))
            {
                //Read length
                var ws = p.Workbook.Worksheets.Add("WIDTH");
                ExcelTextFormatFixedWidth fw = new ExcelTextFormatFixedWidth();
                int[] arr = { 8, 4, 11, 13, 27, 5, 5, 9, 4, 18, 20, 8 };
                fw.SetColumns(FixedWidthReadType.Length, arr);
                fw.SkipLinesBeginning = 1;
                fw.ShouldUseRow = row =>
                {
                    if (row.Length >= fw.LineLength)
                    {
                        if (row.Contains("Page") || string.IsNullOrEmpty(row))
                        {
                            return false;
                        }
                    }
                    return true;
                };
                ws.Cells["A1"].LoadFromText(myFileInfo, fw, TableStyles.Dark10, true);

                //Read positions
                var ws2 = p.Workbook.Worksheets.Add("POSITION");
                ExcelTextFormatFixedWidth fw2 = new ExcelTextFormatFixedWidth();
                fw2.ReadType = FixedWidthReadType.Positions;
                fw2.SetColumnPositions( 0, 8, 12, 24, 35, 63, 68, 73, 82, 86, 105, 125 );
                fw2.SkipLinesBeginning = 1;
                fw2.ShouldUseRow = row =>
                {
                    if (row.Length >= fw2.LineLength)
                    {
                        if (row.Contains("Page") || string.IsNullOrEmpty(row))
                        {
                            return false;
                        }
                    }
                    return true;
                };
                ws2.Cells["A1"].LoadFromText(myFileInfo, fw2, TableStyles.Dark10, true);


                //Read widths 2 Cols
                var ws3 = p.Workbook.Worksheets.Add("WIDTH3");
                ExcelTextFormatFixedWidth fw3 = new ExcelTextFormatFixedWidth();
                fw3.SetColumnLengths(8, 4, 11, 13, 27, 5, 5, 9, 4, 18, 20, 8);
                fw3.SetUseColumns( false, false, true, true, false, false, false, false, false, true, true, false );
                fw3.SkipLinesBeginning = 1;
                ws3.Cells["A1"].LoadFromText(myFileInfo, fw3, TableStyles.Medium5, true);
                //Read positions 3 cols
                var ws4 = p.Workbook.Worksheets.Add("POSITION2");
                ExcelTextFormatFixedWidth fw4 = new ExcelTextFormatFixedWidth();
                fw4.ReadType = FixedWidthReadType.Positions;
                fw4.SetColumnPositions( 0, 8, 12, 24, 35, 63, 68, 73, 82, 86, 105, 125 );
                fw4.SetUseColumns( false, false, true, true, false, false, false, false, false, true, true, false );
                fw4.SkipLinesBeginning = 1;
                ws4.Cells["A1"].LoadFromText(myFileInfo, fw4, TableStyles.Medium5, true);
                p.Save();
            }
        }

        [TestMethod]
        public void ReadFileFromDiskAndSaveAsXLSXFixedWidth2()
        {
            string myFile = File.ReadAllText("C:\\Users\\AdrianParnéus\\Documents\\Test\\FW2.txt");
            FileInfo myFileInfo = new FileInfo("C:\\Users\\AdrianParnéus\\Documents\\Test\\FW2.txt");
            using (var p = OpenTemplatePackage("Fixed2.xlsx"))
            {
                //Read length
                var ws = p.Workbook.Worksheets.Add("TEST");
                ExcelTextFormatFixedWidth fw = new ExcelTextFormatFixedWidth();
                fw.SkipLinesBeginning = 36;
                fw.SkipLinesEnd = 6;
                int[] arr = { 0, 16, 32, 43, 55, 62 };
                fw.SetColumns(FixedWidthReadType.Positions, arr);
                fw.SetColumnsNames("Name", "Position", "Prot", "Entry_Name", "Code", "Description");
                ws.Cells["A1"].LoadFromText(myFileInfo, fw, TableStyles.Dark10, true);
                p.Save();
            }
        }

        [TestMethod]
        public void ReadFixedTextWidthExample()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Name            Date      Ammount          Percent Category");
            sb.AppendLine("David           2024/03/02          130000      2% A");
            sb.AppendLine("Meryl           2024/02/15             999     10% B");
            sb.AppendLine("Hal             2005/11/24               0      0% A");
            sb.AppendLine("Frank           1988/10/12           40,00     59% C");
            sb.AppendLine("Naomi           2015/09/03       245000,99    100% C");
            string myFile = sb.ToString();


            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();
                format.SetColumnPositions(0, 16, 26, 42, 50);
                format.ReadType = FixedWidthReadType.Positions;
                format.SetColumnLengths(16, 10, 16, 8, 2);
                format.SetColumnPaddingAlignmentType(PaddingAlignmentType.Left, PaddingAlignmentType.Auto, PaddingAlignmentType.Right, PaddingAlignmentType.Right, PaddingAlignmentType.Auto);
                format.SetColumnDataTypes(eDataTypes.String, eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Percent, eDataTypes.String);
                ws.Cells["A1"].LoadFromText(myFile, format);
                
                Assert.AreEqual("David", ws.Cells["A2"].Value);
                Assert.AreEqual("C", ws.Cells["E6"].Value);
            }

        }

        [TestMethod]
        public void ReadFixedTextWidthTranposed()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Name            Date      Ammount          Percent Category");
            sb.AppendLine("David           2024/03/02          130000      2% A");
            sb.AppendLine("Meryl           2024/02/15             999     10% B");
            sb.AppendLine("Hal             2005/11/24               0      0% A");
            sb.AppendLine("Frank           1988/10/12           40,00     59% C");
            sb.AppendLine("Naomi           2015/09/03       245000,99    100% C");
            string myFile = sb.ToString();


            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();
                format.SetColumnPositions(0, 16, 26, 42, 50);
                format.ReadType = FixedWidthReadType.Positions;
                format.SetColumnLengths(16, 10, 16, 8, 2);
                format.Transpose = true;
                var r = ws.Cells["A1"].LoadFromText(myFile, format);
                Assert.AreEqual("A1:F5", r.Address);
                Assert.AreEqual("David", ws.Cells["B1"].Value);
                Assert.AreEqual("C", ws.Cells["F5"].Value);
            }

        }

    }
}
