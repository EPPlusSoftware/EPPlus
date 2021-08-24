using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    interface IExcelColumn
    {
        /// <summary>
        /// If the column is collapsed in outline mode
        /// </summary>
        bool Collapsed { get; set; }
        /// <summary>
        /// Outline level. Zero if no outline
        /// </summary>
        int OutlineLevel { get; set; }
        /// <summary>
        /// Phonetic
        /// </summary>
        bool Phonetic { get; set; }
        /// <summary>
        /// If set to true a column automaticlly resize(grow wider) when a user inputs numbers in a cell. 
        /// </summary>
        bool BestFit
        {
            get;
            set;
        }
        void AutoFit();
        void AutoFit(double MinimumWidth);
        /// <summary>
        /// Set the column width from the content.
        /// Note: Cells containing formulas are ignored unless a calculation is performed.
        ///       Wrapped and merged cells are also ignored.
        /// </summary>
        /// <param name="MinimumWidth">Minimum column width</param>
        /// <param name="MaximumWidth">Maximum column width</param>
        void AutoFit(double MinimumWidth, double MaximumWidth);
        bool Hidden
        {
            get;
            set;
        }
        double Width
        {
            get;
            set;
        }
        /// <summary>
        /// Adds a manual page break after the column.
        /// </summary>
        bool PageBreak
        {
            get;
            set;
        }
        /// <summary>
        /// Merges all cells of the column
        /// </summary>
        bool Merged
        {
            get;
            set;
        }
    }
    public class ExcelRangeColumn : IExcelColumn, IEnumerable<ExcelRangeColumn>, IEnumerator<ExcelRangeColumn>
    {
        ExcelWorksheet _worksheet;
        internal int _fromCol, _toCol;
        internal ExcelRangeColumn(ExcelWorksheet ws, int fromCol, int toCol)
        {
            _worksheet = ws;
            _fromCol = fromCol;
            _toCol = toCol;            
        }
        /// <summary>
        /// The first column in the collection
        /// </summary>
        public int StartColumn 
        { 
            get
            {
                return _fromCol;
            }
        }
        /// <summary>
        /// The last column in the collection
        /// </summary>
        public int EndColumn
        {
            get
            {
                return _toCol;
            }
        }
        public bool Collapsed 
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.Collapsed), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.Collapsed = v; }), value);
            }
        }
        public int OutlineLevel
        {
            get
            {
                return GetValue(new Func<ExcelColumn, int>(x => x.OutlineLevel), 0);
            }
            set
            {
                SetValue(new Action<ExcelColumn, int>((x, v) => { x.OutlineLevel = v; }), value);
            }
        }

        public bool Phonetic
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.Phonetic), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.Phonetic = v; }), value);
            }
        }
        public bool BestFit
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.BestFit), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.BestFit = v; }), value);
            }
        }

        public bool Hidden
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.Hidden), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.Hidden = v; }), value);
            }
        }
        public double Width
        {
            get
            {
                return GetValue(new Func<ExcelColumn, double>(x => x.Width), _worksheet.DefaultColWidth);
            }
            set
            {
                SetValue(new Action<ExcelColumn, double>((x, v) => { x.Width = v; }), value);
            }
        }

        /// <summary>
        /// Adds a manual page break after the column.
        /// </summary>
        public bool PageBreak
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.PageBreak), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.PageBreak = v; }), value);
            }
        }
        /// <summary>
        /// Merges all cells of the column
        /// </summary>
        public bool Merged
        {
            get
            {
                return GetValue(new Func<ExcelColumn, bool>(x => x.Merged), false);
            }
            set
            {
                SetValue(new Action<ExcelColumn, bool>((x, v) => { x.Merged = v; }), value);
            }
        }
        #region ExcelColumn Style
        /// <summary>
        /// The Style applied to the whole column(s). Only effects cells with no individual style set. 
        /// Use Range object if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                string letter = ExcelCellBase.GetColumnLetter(_fromCol);
                string endLetter = ExcelCellBase.GetColumnLetter(_toCol);
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionId, letter + ":" + endLetter);
            }
        }
        internal string _styleName = "";
        /// <summary>
		/// Sets the style for the entire column using a style name.
		/// </summary>
		public string StyleName
        {

            get
            {
                return GetValue<string>(new Func<ExcelColumn, string>(x => x.StyleName), "");
            }
            set
            {
                SetValue(new Action<ExcelColumn,string>((x,v) => { x.StyleName = v; }), value);
            }
        }
        /// <summary>
        /// Sets the style for the entire column using the style ID.           
        /// </summary>
        public int StyleID
        {
            get
            {
                return GetValue(new Func<ExcelColumn, int>(x => x.StyleID), 0);
            }
            set
            {
                SetValue(new Action<ExcelColumn, int>((x, v) => { x.StyleID = v; }), value);
            }
        }

        public ExcelRangeColumn Current
        {
            get
            {
                return new ExcelRangeColumn(_worksheet, enumCol, enumCol);
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return new ExcelRangeColumn(_worksheet, enumCol, enumCol);
            }
        }
        #endregion

        public void AutoFit()
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns();
        }

        public void AutoFit(double MinimumWidth)
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth);
        }

        public void AutoFit(double MinimumWidth, double MaximumWidth)
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth, MaximumWidth);
        }
        private TOut GetValue<TOut>(Func<ExcelColumn, TOut> getValue, TOut defaultValue)
        {
            var currentCol = _worksheet.GetValueInner(0, _fromCol) as ExcelColumn;
            if (currentCol == null)
            {
                int r = 0, c = _fromCol;
                if(_worksheet._values.PrevCell(ref r, ref c))
                {
                    if(c>0)
                    {
                        ExcelColumn prevCol = _worksheet.GetValueInner(0, c) as ExcelColumn;
                        if (prevCol.ColumnMax>=_fromCol)
                        {
                            return getValue(prevCol);
                        }
                    }
                }
                return defaultValue;
            }
            else
            {
                return getValue(currentCol);
            }
        }

        private void SetValue<T>(Action<ExcelColumn,T> SetValue, T value)
        {
            var c = _fromCol;
            int r = 0;
            ExcelColumn currentCol = _worksheet.GetValueInner(0, c) as ExcelColumn;
            while (c <= _toCol)
            {
                if (currentCol == null)
                {
                    currentCol = _worksheet.Column(c);
                }
                else
                {
                    if (c < _fromCol)
                    {
                        currentCol = _worksheet.Column(c);
                    }
                }

                if (currentCol.ColumnMax >= _toCol)
                {
                    currentCol.ColumnMax = _toCol;
                }
                else
                {
                    if (_worksheet._values.NextCell(ref r, ref c))
                    {
                        if (r == 0 && c < _toCol)
                        {
                            currentCol.ColumnMax = c - 1;
                        }
                        else
                        {
                            currentCol.ColumnMax = _toCol;
                        }
                    }
                    else
                    {
                        currentCol.ColumnMax = _toCol;
                    }
                }
                c = currentCol.ColumnMax + 1;
                SetValue(currentCol, value);

            }
        }

        public IEnumerator<ExcelRangeColumn> GetEnumerator()
        {
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }

        public bool MoveNext()
        {
            if(_cs==null)
            {
                Reset();
                return enumCol <= _toCol;
            }
            enumCol++;
            if (_currentCol?.ColumnMax>=enumCol)
            {
                return true;
            }
            else
            {
                var c = _cs.GetValue(0, enumCol)._value as ExcelColumn;
                if(c!=null && c.ColumnMax>=enumCol)
                {
                    _currentCol = c;
                    return true;
                }
                if(++enumColPos<_cs.ColumnCount)
                {
                    enumCol = _cs._columnIndex[enumColPos].Index;
                }
                else
                {
                    return false;
                }
                if (enumCol <= _toCol) 
                    return true;
            }
            return false;
        }
        CellStoreValue _cs;
        int enumCol, enumColPos;
        ExcelColumn _currentCol;
        public void Reset()
        {
            _currentCol = null;
            _cs = _worksheet._values;
            if(_cs.ColumnCount>0)
            {
                enumCol = _fromCol;
                enumColPos = _cs.GetColumnPosition(enumCol);
                if(enumColPos<0)
                {
                    enumColPos = ~enumColPos;
                    int r=0, c=0;
                    if(_cs.GetPrevCell(ref r, ref c, 0, enumColPos - 1, _toCol))
                    {
                        if (r == 0 && c < enumColPos)
                        {
                            _currentCol = ((ExcelColumn)_cs.GetValue(r, c)._value);
                            if (_currentCol.ColumnMax >= _fromCol)
                            {
                                enumColPos = c;
                            }
                            else
                            {
                                enumCol = _cs._columnIndex[enumColPos].Index;
                            }
                        }
                    }
                    else
                    {
                        enumColPos++;
                        enumCol = _cs._columnIndex[enumColPos].Index;
                    }

                }
                else
                {
                    _currentCol = _cs.GetValue(0, _fromCol)._value as ExcelColumn; 
                }
            }
        }

        public void Dispose()
        {
            
        }
    }
}
