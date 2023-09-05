using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    [DebuggerDisplay("Value: {Value}")]
    internal class InMemoryCellInfo : ICellInfo
    {
        public InMemoryCellInfo(object value)
        {
            _value = value;
        }

        private readonly object _value;
        public string Address => null;

        public string WorksheetName => null;

        public int Row => 0;

        public int Column => 0;

        public ulong Id => 0;

        public string Formula => String.Empty;

        public object Value => _value;

        public double ValueDouble => ConvertUtil.GetValueDouble(_value, true);

        public double ValueDoubleLogical => ConvertUtil.GetValueDouble(_value, false);

        public bool IsHiddenRow => false;

        public bool IsExcelError => ExcelErrorValue.IsErrorValue(_value?.ToString());

        public IList<Token> Tokens => throw new NotImplementedException();
    }
}
