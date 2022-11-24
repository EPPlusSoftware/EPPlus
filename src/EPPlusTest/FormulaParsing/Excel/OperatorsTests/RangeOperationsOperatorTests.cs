using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.OperatorsTests
{

    [TestClass]
    public class RangeOperationsOperatorTests
    {
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _context = ParsingContext.Create(_package);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        private ParsingContext _context;
        private ExcelPackage _package;

        [TestMethod]
        public void ShouldSetNAerrorWithDifferentColSize()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);
            r2.SetValue(1, 0, 2);
            r2.SetValue(1, 1, 3);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(5d, range.GetValue(0, 1));
            Assert.AreEqual(3d, range.GetValue(1, 0));
            Assert.AreEqual(5d, range.GetValue(1, 1));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(0, 2));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(1, 2));
        }

        [TestMethod]
        public void ShouldSetNAerrorWithDifferentRowSize()
        {
            var rd1 = new RangeDefinition(3, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            r1.SetValue(2, 0, 1);
            r1.SetValue(2, 1, 2);
            r1.SetValue(2, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 3);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 2);
            r2.SetValue(0, 2, 3);
            r2.SetValue(1, 0, 1);
            r2.SetValue(1, 1, 2);
            r2.SetValue(1, 2, 3);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(4d, range.GetValue(0, 1));
            Assert.AreEqual(6d, range.GetValue(0, 2));
            Assert.AreEqual(2d, range.GetValue(1, 0));
            Assert.AreEqual(4d, range.GetValue(1, 1));
            Assert.AreEqual(6d, range.GetValue(1, 2));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(2, 0));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(2, 1));
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), range.GetValue(2, 2));
        }

        [TestMethod]
        public void ShouldCalculateWithSameRowSize()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 1);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(1, 0, 2);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(3d, range.GetValue(1, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(4d, range.GetValue(1, 1));
            Assert.AreEqual(4d, range.GetValue(0, 2));
            Assert.AreEqual(5d, range.GetValue(1, 2));
        }

        [TestMethod]
        public void ShouldCalculateWithSameColumnSize()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(2d, range.GetValue(1, 0));
            Assert.AreEqual(5d, range.GetValue(0, 1));
            Assert.AreEqual(5d, range.GetValue(1, 1));
        }

        [TestMethod]
        public void ShouldCalculateWithRangeAndSingleCell()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 1);
            var r2 = new InMemoryRange(rd2);
            r2.SetValue(0, 0, 1);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(2d, range.GetValue(1, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(3d, range.GetValue(1, 1));
            Assert.AreEqual(4d, range.GetValue(0, 2));
            Assert.AreEqual(4d, range.GetValue(1, 2));
        }

        [TestMethod]
        public void ShouldCalculateWithRangeAndSingleNumberRight()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);
            var c2 = new CompileResult(1, DataType.Integer);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(2d, range.GetValue(1, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(3d, range.GetValue(1, 1));
            Assert.AreEqual(4d, range.GetValue(0, 2));
            Assert.AreEqual(4d, range.GetValue(1, 2));
        }

        [TestMethod]
        public void ShouldCalculateWithRangeAndSingleNumberLeft()
        {
            var rd1 = new RangeDefinition(2, 3);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            r1.SetValue(0, 2, 3);
            r1.SetValue(1, 0, 1);
            r1.SetValue(1, 1, 2);
            r1.SetValue(1, 2, 3);
            var c1 = new CompileResult(r1, DataType.ExcelRange);
            var c2 = new CompileResult(1, DataType.Integer);

            var result = RangeOperationsOperator.Apply(c2, c1, Operators.Plus, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(2d, range.GetValue(0, 0));
            Assert.AreEqual(2d, range.GetValue(1, 0));
            Assert.AreEqual(3d, range.GetValue(0, 1));
            Assert.AreEqual(3d, range.GetValue(1, 1));
            Assert.AreEqual(4d, range.GetValue(0, 2));
            Assert.AreEqual(4d, range.GetValue(1, 2));
        }

        [TestMethod]
        public void ShouldCalculateRangesNumericWithEqualOperator()
        {
            var rd1 = new RangeDefinition(1, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c2, c1, Operators.Equals, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0));
            Assert.AreEqual(false, range.GetValue(0, 1));
        }

        [TestMethod]
        public void ShouldCalculateRangesStringWithEqualOperator()
        {
            var rd1 = new RangeDefinition(1, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, "A");
            r1.SetValue(0, 1, "b");
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, "a");
            r2.SetValue(0, 1, "C");

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c2, c1, Operators.Equals, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0));
            Assert.AreEqual(false, range.GetValue(0, 1));
        }

        [TestMethod]
        public void ShouldCalculateRangesNumericWithNotEqualOperator()
        {
            var rd1 = new RangeDefinition(1, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1);
            r1.SetValue(0, 1, 2);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 1);
            r2.SetValue(0, 1, 3);

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c2, c1, Operators.NotEqualTo, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(false, range.GetValue(0, 0));
            Assert.AreEqual(true, range.GetValue(0, 1));
        }

        [TestMethod]
        public void ShouldCalculateRangesStringWithNotEqualOperator()
        {
            var rd1 = new RangeDefinition(1, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, "A");
            r1.SetValue(0, 1, "b");
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(1, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, "a");
            r2.SetValue(0, 1, "C");

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c2, c1, Operators.NotEqualTo, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(false, range.GetValue(0, 0));
            Assert.AreEqual(true, range.GetValue(0, 1));
        }

        [TestMethod]
        public void ShouldCalculateRangesStringWithLessThanOrEqualOperator()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, "Abc");
            r1.SetValue(0, 1, "b");
            r1.SetValue(1, 0, "basdf");
            r1.SetValue(1, 1, "b");
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, "abc");
            r2.SetValue(0, 1, "C");
            r2.SetValue(1, 0, "a");
            r2.SetValue(1, 1, "C");

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.LessThanOrEqual, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0), "'Abc' was not less than or equal to 'abc'");
            Assert.AreEqual(true, range.GetValue(0, 1), "'b' was not less than or equal to 'C'");
            Assert.AreEqual(false, range.GetValue(1, 0), "'basdf' was not less than or equal to 'a'");
        }

        [TestMethod]
        public void ShouldCalculateRangesDoubleWithLessThanOrEqualOperator()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1.1d);
            r1.SetValue(0, 1, 1d);
            r1.SetValue(1, 0, 2d);
            r1.SetValue(1, 1, 3.001d);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 1.1d);
            r2.SetValue(0, 1, 2d);
            r2.SetValue(1, 0, 1d);
            r2.SetValue(1, 1, 3d);

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.LessThanOrEqual, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0), "'1.1' was not less than or equal to '1.1'");
            Assert.AreEqual(true, range.GetValue(0, 1), "'1' was not less than or equal to '2'");
            Assert.AreEqual(false, range.GetValue(1, 0), "'2' was considered less than or equal to '1'");
        }

        [TestMethod]
        public void ShouldCalculateRangesStringWithGreaterThanOrEqualThanOperator()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, "Abc");
            r1.SetValue(0, 1, "b");
            r1.SetValue(1, 0, "basdf");
            r1.SetValue(1, 1, "b");
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, "abc");
            r2.SetValue(0, 1, "C");
            r2.SetValue(1, 0, "a");
            r2.SetValue(1, 1, "C");

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.GreaterThanOrEqual, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0), "'Abc' was not less than or equal to 'abc'");
            Assert.AreEqual(false, range.GetValue(0, 1), "'b' was not less than or equal to 'C'");
            Assert.AreEqual(true, range.GetValue(1, 0), "'basdf' was not less than or equal to 'a'");
        }

        [TestMethod]
        public void ShouldCalculateRangesDoubleWithGreaterThanOrEqualOperator()
        {
            var rd1 = new RangeDefinition(2, 2);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1.1d);
            r1.SetValue(0, 1, 1d);
            r1.SetValue(1, 0, 2d);
            r1.SetValue(1, 1, 3.001d);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 2);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 1.1d);
            r2.SetValue(0, 1, 2d);
            r2.SetValue(1, 0, 1d);
            r2.SetValue(1, 1, 3d);

            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.GreaterThanOrEqual, _context);
            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(true, range.GetValue(0, 0), "'1.1' was not greater than or equal to '1.1'");
            Assert.AreEqual(false, range.GetValue(0, 1), "'1' was not greater than or equal to '2'");
            Assert.AreEqual(true, range.GetValue(1, 0), "'2' was considered greater than or equal to '1'");
        }

        [TestMethod]
        public void ShouldCalculateRangesDoubleWithConcatenateOperator()
        {
            var rd1 = new RangeDefinition(2, 1);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 1d);
            r1.SetValue(1, 0, 2d);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 1);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 3d);
            r2.SetValue(1, 0, 4d);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Concat, _context);

            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual("13", range.GetValue(0, 0));
            Assert.AreEqual("24", range.GetValue(1, 0));
        }

        [TestMethod]
        public void ShouldCalculateRangesStringWithConcatenateOperator()
        {
            var rd1 = new RangeDefinition(2, 1);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, "a");
            r1.SetValue(1, 0, "b");
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 1);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, "c");
            r2.SetValue(1, 0, "d");
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Concat, _context);

            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual("ac", range.GetValue(0, 0));
            Assert.AreEqual("bd", range.GetValue(1, 0));
        }

        [TestMethod]
        public void ShouldCalculateRangesDoubleWithExpOperator()
        {
            var rd1 = new RangeDefinition(2, 1);
            var r1 = new InMemoryRange(rd1);
            r1.SetValue(0, 0, 10d);
            r1.SetValue(1, 0, 2d);
            var c1 = new CompileResult(r1, DataType.ExcelRange);

            var rd2 = new RangeDefinition(2, 1);
            var r2 = new InMemoryRange(rd1);
            r2.SetValue(0, 0, 3d);
            r2.SetValue(1, 0, 5d);
            var c2 = new CompileResult(r2, DataType.ExcelRange);

            var result = RangeOperationsOperator.Apply(c1, c2, Operators.Exponentiation, _context);

            Assert.IsInstanceOfType(result.ResultValue, typeof(InMemoryRange));
            var range = result.ResultValue as InMemoryRange;
            Assert.IsNotNull(range);
            Assert.AreEqual(1000d, range.GetValue(0, 0));
            Assert.AreEqual(32d, range.GetValue(1, 0));
        }
    }
}
