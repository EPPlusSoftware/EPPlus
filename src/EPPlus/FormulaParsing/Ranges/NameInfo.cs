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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{    
    /// <summary>
    /// EPPlus implementation of the <see cref="INameInfo"/> interface
    /// </summary>
    public class NameInfo : INameInfo
    {
        ExcelNamedRange _nameItem;
        /// <summary>
        /// Nameinfo
        /// </summary>
        /// <param name="nameItem"></param>
        public NameInfo(ExcelNamedRange nameItem)
        {
            _nameItem=nameItem;
        }
        /// <summary>
        /// Id
        /// </summary>
        public ulong Id 
        {
            get
            {
                if(_nameItem==null)
                {
                    return ulong.MaxValue;
                }
                return ExcelCellBase.GetCellId(_nameItem.LocalSheetId, _nameItem.Index, 0);
            }
        }
        /// <summary>
        /// Worksheet name
        /// </summary>
        public int wsIx 
        {
            get
            {
                return (_nameItem?.Worksheet == null ? int.MinValue : _nameItem.Worksheet.IndexInList);
            }
        }
        /// <summary>
        /// The name
        /// </summary>
        public string Name
        {
            get
            {
                return _nameItem.Name;
            }
        }
        /// <summary>
        /// Formula of the name
        /// </summary>
        public string Formula
        {
            get
            {
                return _nameItem.Formula;
            }
        }
        /// <summary>
        /// Gets the forumla relative to a row and column.
        /// </summary>
        /// <param name="row">The row </param>
        /// <param name="col">The column</param>
        /// <returns></returns>
        public string GetRelativeFormula(int row, int col)
        {
            return _nameItem.GetRelativeFormula(row, col);
        }
        /// <summary>
        /// Returns the range relative to the cell for a named range with a relative address.
        /// </summary>
        /// <param name="ri"></param>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        public IRangeInfo GetRelativeRange(IRangeInfo ri, FormulaCellAddress currentCell)
        {
            return _nameItem.GetRelativeRange(ri, currentCell);
        }

        /// <summary>
        /// Get the value relative to the current cell.
        /// </summary>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        public object GetValue(FormulaCellAddress currentCell)
        {
            return _nameItem.GetValue(currentCell);
        }
        /// <summary>
        /// 
        /// </summary>
        public bool IsRelative
        {
            get
            {
                return _nameItem.IsRelative;
            }
        }
        /// <summary>
        /// Tokens
        /// </summary>
        public IList<Token> Tokens { get; internal set; }
        /// <summary>
        /// Value
        /// </summary>
        public object Value
        {
            get
            {
                return _nameItem.NameValue;
            }
            set
            {
                _nameItem.NameValue = value;
            }
        }
        
    }
    /// <summary>
    /// Name info with value
    /// </summary>
    public class NameInfoWithValue : INameInfo
    {
        string _name;
        /// <summary>
        /// Name info with value
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public NameInfoWithValue(string name, object value)
        {
            _name = name;
            Value = value;
        }
        /// <summary>
        /// Id
        /// </summary>
        public ulong Id => long.MaxValue;
        /// <summary>
        /// wsIx
        /// </summary>
        public int wsIx => -1;
        /// <summary>
        /// Name
        /// </summary>
        public string Name => _name;
        /// <summary>
        /// Formula
        /// </summary>
        public string Formula => "";
        /// <summary>
        /// Value
        /// </summary>
        public object Value 
        {
            get;
            private set;
        }
        /// <summary>
        /// IsRelative
        /// </summary>
        public bool IsRelative => false;

        /// <summary>
        /// GetValue
        /// </summary>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        public object GetValue(FormulaCellAddress currentCell)
        {
            return Value;
        }
        /// <summary>
        /// Get relative formula
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public string GetRelativeFormula(int row, int col)
        {
            return null;
        }
        /// <summary>
        /// Get relative range
        /// </summary>
        /// <param name="ri"></param>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        public IRangeInfo GetRelativeRange(IRangeInfo ri, FormulaCellAddress currentCell)
        {
            return null;
        }
    }
}
