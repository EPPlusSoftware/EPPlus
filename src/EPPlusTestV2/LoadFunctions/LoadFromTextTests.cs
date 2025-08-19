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
using System.Globalization;
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
            _formatFixed.SetColumnLengths( 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
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
            _formatFixed.SetColumnLengths( 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _formatFixed.Columns[0].DataType = eDataTypes.Number;
            _formatFixed.Columns[1].DataType = eDataTypes.Number;
            _formatFixed.Columns[2].DataType = eDataTypes.DateTime;
            _formatFixed.Columns[3].DataType = eDataTypes.Number;
            _formatFixed.Columns[4].DataType = eDataTypes.String;
            _formatFixed.Columns[5].DataType = eDataTypes.String;
            _formatFixed.Columns[6].DataType = eDataTypes.String;
            _formatFixed.Columns[7].DataType = eDataTypes.String;
            _formatFixed.Columns[8].DataType = eDataTypes.String;
            _formatFixed.Columns[9].DataType = eDataTypes.Number;
            _formatFixed.Columns[10].DataType = eDataTypes.Number;
            _formatFixed.Columns[11].DataType = eDataTypes.String;
            _formatFixed.TableStyle = TableStyles.None;
            _formatFixed.FirstRowIsHeader = true;
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

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
            _formatFixed.SetColumnLengths(8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _formatFixed.TableStyle = TableStyles.None;
            _formatFixed.FirstRowIsHeader = true;
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(16524d, _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldReturnRangeFixedWidth()
        {
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            _formatFixed.SetColumnLengths(8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
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
            _formatFixed.SetColumnLengths(8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
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
            _formatFixed.SetColumnLengths( 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8);
            _formatFixed.Columns[0].UseColumn = false;
            _formatFixed.Columns[1].UseColumn = false;
            _formatFixed.Columns[2].UseColumn = true;
            _formatFixed.Columns[3].UseColumn = true;
            _formatFixed.Columns[4].UseColumn = false;
            _formatFixed.Columns[5].UseColumn = false;
            _formatFixed.Columns[6].UseColumn = false;
            _formatFixed.Columns[7].UseColumn = false;
            _formatFixed.Columns[8].UseColumn = false;
            _formatFixed.Columns[9].UseColumn = false;
            _formatFixed.Columns[10].UseColumn = true;
            _formatFixed.Columns[11].UseColumn = false;
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
            _formatFixed.SetColumnPositions(0, arr );
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(2368183100d, _worksheet.Cells["D3"].Value);
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
                format.SetColumnPositions(0, 0, 16, 26, 42, 50);
                format.ReadType = FixedWidthReadType.Positions;
                //format.SetColumnLengths(16, 10, 16, 8, 2);
                format.Transpose = true;
                var r = ws.Cells["A1"].LoadFromText(myFile, format);
                Assert.AreEqual("A1:F5", r.Address);
                Assert.AreEqual("David", ws.Cells["B1"].Value);
                Assert.AreEqual("C", ws.Cells["F5"].Value);
            }

        }

        [TestMethod]
        public void ReadFixedTextWidthTrailingMinus()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Name            Date      Ammount          Percent Category");
            sb.AppendLine("David           2024/03/02         130000-      2% A");
            sb.AppendLine("Meryl           2024/02/15             999     10% B");
            sb.AppendLine("Hal             2005/11/24               0      0% A");
            sb.AppendLine("Frank           1988/10/12          40,00-     59% C");
            sb.AppendLine("Naomi           2015/09/03       245000,99    100% C");
            string myFile = sb.ToString();


            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();
                format.SetColumnPositions(52, 0, 16, 26, 42, 50);
                //format.SetColumnPaddingAlignmentType(PaddingAlignmentType.Left, PaddingAlignmentType.Auto, PaddingAlignmentType.Right, PaddingAlignmentType.Right, PaddingAlignmentType.Auto);
                //format.SetColumnDataTypes(eDataTypes.String, eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Percent, eDataTypes.String);
                format.Culture = new CultureInfo("sv-en");
                var range = ws.Cells["A1"].LoadFromText(myFile, format);
                Assert.AreEqual(-130000d, ws.Cells["C2"].Value);
                Assert.AreEqual(-40.00d, ws.Cells["C5"].Value);
            }
        }
        [TestMethod]
        public void ReadFixedTextFileList()
        {
            var fileContent = Properties.Resources.GetTextFileContent("FixedWidth_FileList.txt", Encoding.GetEncoding(437));
            if (string.IsNullOrEmpty(fileContent))
            {
                var fi = Properties.Resources.GetTextFileInfo("FixedWidth_FileList.txt");
                if(fi.Exists==false)
                {
                    Assert.Fail("File does not exist : " + fi.FullName);
                }
                else
                {
                    Assert.Fail("File Content is empty for file " + fi.FullName);
                }
            }

            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();

                format.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
                format.SetColumnLengths(12, 9, 5, 10, -1);
                format.SkipLinesBeginning = 5;
                format.SkipLinesEnd= 2;
                format.Culture = CultureInfo.GetCultureInfo("sv-en");
                format.TableStyle = TableStyles.Medium12;
                format.SetColumnsNames("Date", "Time", "Type","Size", "Name");
                format.EOL = "\n";
                var range = ws.Cells["A1"].LoadFromText(fileContent, format);
                if(range.Rows<101)
                {
                    Assert.Fail($"Load failed. LoadRange is {range.Address}, {range.Offset(0,0, 1, 1).Value}, {range.Offset(0, 1 ,1, 1).Value}, {range.Offset(0, 2, 1, 1).Value}, {range.Offset(0, 3, 1, 1).Value}, {range.Offset(0, 4).Value}");
                }

                range.TakeSingleColumn(0).Style.Numberformat.Format = "yyyy-MM-dd";
                range.TakeSingleColumn(1).Style.Numberformat.Format = "hh:mm";
                range.TakeSingleColumn(2).Style.Numberformat.Format = "#,##0";
                ws.Cells.AutoFitColumns();

                Assert.AreEqual(new DateTime(2022, 11, 22), ws.Cells["A14"].Value);
                Assert.AreEqual(30, ((DateTime)ws.Cells["B14"].Value).Minute);
                Assert.AreEqual(6136D, ws.Cells["D14"].Value);
                Assert.AreEqual("Enums.cs", ws.Cells["E14"].Value);

                SaveWorkbook("FileList.xlsx",p);
            }
        }
        [TestMethod]
        public void ReadFixedTextFile2()
        {
            var file = Properties.Resources.GetTextFileInfo("FixedWidth1.txt");
            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();

                format.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
                format.SetColumnPositions(-1, 0, 30, 60, 80);
                format.Culture = CultureInfo.InvariantCulture;
                format.TableStyle = TableStyles.Medium2;
                format.FirstRowIsHeader = true;
                format.Encoding = Encoding.GetEncoding(437);

                var range = ws.Cells["A1"].LoadFromText(file, format);
                ws.Cells.AutoFitColumns();


                SaveWorkbook("FixedWidth1.xlsx", p);
            }
        }

        [TestMethod]
        public void ReadFixedTextFile3()
        {
            var file = Properties.Resources.GetTextFileInfo("FixedWidth2.txt");
            //Do fixed width text stuff
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();

                format.FormatErrorStrategy = FixedWidthFormatErrorStrategy.Truncate;
                format.SetColumnLengths(15);
                format.SetColumnPositions(-1, 40);
                format.Culture = CultureInfo.InvariantCulture;
                format.TableStyle = TableStyles.Medium2;
                format.FirstRowIsHeader = true;
                format.Encoding = Encoding.GetEncoding(437);

                var range = ws.Cells["A1"].LoadFromText(file, format);
                ws.Cells.AutoFitColumns();


                SaveWorkbook("FixedWidth2.xlsx", p);
            }
        }

    }
}
