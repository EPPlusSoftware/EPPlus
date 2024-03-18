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
        public void ShouldLoadFixedWidthText()
        {   //          6     6        10       12          28                      5       6     6     6                  19                   20       8         
            AddLine(" Entry  Per. Post Date  GL Account   Description               Srce. Cflow  Ref.      Post             Debit              Credit  Alloc.");
            AddLine(" 16524  01  10/17/2012  3930621977   TXNPUES                   S1    Yes    RHMXWPCP  Yes                               5,007.10  No  ");
            AddLine("191675  01  01/14/2013  2368183100   OUNHQEX XUFQONY           S1    No               Yes                              43,537.00  Yes ");
            AddLine("191667  01  01/14/2013  3714468136   GHAKASC QHJXDFM           S1    Yes              Yes           3,172.53                      Yes ");

            using(var p = new ExcelPackage("C:\\epplusTest\\Testoutput\\FixedWidth.xlsx") )
            {

                var ws = p.Workbook.Worksheets.Add("SHHET");
                ws.Cells["A1"].LoadFromFixedWidthText(_lines.ToString(), 8, 4, 12, 13, 26, 6, 6, 6, 6, 19, 20, 8);
                Assert.AreEqual("Entry", _worksheet.Cells["A1"].Value);
                p.Save();
            }
        }
    }
}
