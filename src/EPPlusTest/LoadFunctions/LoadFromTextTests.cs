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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromTextTests
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
        public void ShouldLoadFixedWidthText()
        {
            ///TODODO
            /* read positon
             * Read from file
             * save sheets to file for testing
             * save to text file
             */

            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8 };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldUseTypesFromFormatFixedWidth()
        {
            //          8     5        11       13          32                             6      6     10        4               20                   20       8         
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            AddLine(" 16524  01   10/17/2012 3930621977   TXNPUES                         S1    Yes   RHMXWPCP  Yes                                 5,007.10  No    ");
            AddLine("191675  01   01/14/2013 2368183100   OUNHQEX XUFQONY                 S1    No              Yes                                43,537.00  Yes   ");
            AddLine("191667  01   01/14/2013 3714468136   GHAKASC QHJXDFM                 S1    Yes             Yes             3,172.53                      Yes   ");
            _format.DataTypes = new eDataTypes[] { eDataTypes.Number, eDataTypes.Number, eDataTypes.DateTime, eDataTypes.Number, eDataTypes.String, eDataTypes.String, eDataTypes.String, eDataTypes.String, eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.String };
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8 };
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
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8 };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed, TableStyles.None, true);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(16524d, _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldReturnRangeFixedWidth()
        {
            AddLine("Entry   Per. Post Date  GL Account   Description                     Srce. Cflow Ref.      Post               Debit              Credit  Alloc.");
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8 };
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
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8};
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
            _formatFixed.ColumnLengths = new int[] { 8, 5, 11, 13, 32, 6, 6, 10, 4, 20, 20, 8 };
            _formatFixed.UseColumns = new bool[] { false, false, true, true, false, false, false, false, false, false, true, false };
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
            _formatFixed.ReadStartPosition = FixedWidthRead.Positions;
            _formatFixed.ColumnLengths = new int[] { 0, 8, 13, 24, 37, 69, 75, 81, 91, 95, 115, 135 };
            //_formatFixed.UseColumns = new bool[] { false, false, true, true, false, false, false, false, false, false, true, false };
            _worksheet.Cells["A1"].LoadFromText(_lines.ToString(), _formatFixed);

            Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
            Assert.AreEqual(2368183100d, _worksheet.Cells["D3"].Value);
        }

    }
}
