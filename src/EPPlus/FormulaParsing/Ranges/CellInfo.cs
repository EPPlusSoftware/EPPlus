/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/31/2022         EPPlus Software AB           EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// EPPlus implementation of the <see cref="ICellInfo"/> interface.
    /// </summary>
    public class CellInfo : ICellInfo
    {
        ExcelWorksheet _ws;
        CellStoreEnumerator<ExcelValue> _values;
        internal CellInfo(ExcelWorksheet ws, CellStoreEnumerator<ExcelValue> values)
        {
            _ws = ws;
            _values = values;
        }
        public string Address
        {
            get { return _values.CellAddress; }
        }

        public int Row
        {
            get { return _values.Row; }
        }

        public int Column
        {
            get { return _values.Column; }
        }

        public string Formula
        {
            get
            {
                return _ws.GetFormula(_values.Row, _values.Column);
            }
        }

        public object Value
        {
            get
            {
                if (_ws._flags.GetFlagValue(_values.Row, _values.Column, CellFlags.RichText))
                {
                    return _ws.GetRichText(_values.Row, _values.Column, null).Text;
                }
                else
                {
                    return _values.Value._value;
                }
            }
        }

        public double ValueDouble
        {
            get { return ConvertUtil.GetValueDouble(_values.Value._value, true); }
        }
        public double ValueDoubleLogical
        {
            get { return ConvertUtil.GetValueDouble(_values.Value._value, false); }
        }
        public bool IsHiddenRow
        {
            get
            {
                var row = _ws.GetValueInner(_values.Row, 0) as RowInternal;
                if (row != null)
                {
                    return row.Hidden || row.Height == 0;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns true if the cell contains an error
        /// </summary>
        public bool IsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(_values.Value._value); }
        }

        /// <summary>
        /// Tokenized cell content
        /// </summary>
        public IList<Token> Tokens
        {
            get
            {
                return _ws._formulaTokens.GetValue(_values.Row, _values.Column);
            }
        }

        /// <summary>
        /// Cell id
        /// </summary>
        public ulong Id
        {
            get
            {
                return ExcelCellBase.GetCellId(_ws.IndexInList, _values.Row, _values.Column);
            }
        }

        /// <summary>
        /// Name of the worksheet
        /// </summary>
        public string WorksheetName
        {
            get { return _ws.Name; }
        }
    }
}
