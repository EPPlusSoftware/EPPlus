using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    /// <summary>
    /// A column in a worksheet
    /// </summary>
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
        /// Groups the columns using an outline. 
        /// Adds one to <see cref="OutlineLevel" /> for each column if the outline level is less than 8.
        /// </summary>
        void Group();
        /// <summary>
        /// Ungroups the columns from the outline. 
        /// Subtracts one from <see cref="OutlineLevel" /> for each column if the outline level is larger that zero. 
        /// </summary>
        void UnGroup();
        /// <summary>
        /// Collapses and hides the column's children. Children are columns immegetaly to the right or left of the column depending on the <see cref="ExcelWorksheet.OutLineSummaryRight"/>
        /// <paramref name="allLevels">If true, all children will be collapsed and hidden. If false, only the children of the referenced columns are collapsed.</paramref>
        /// </summary>
        void CollapseChildren(bool allLevels = true);
        /// <summary>
        /// Expands and shows the column's children. Children are columns immegetaly to the right or left of the column depending on the <see cref="ExcelWorksheet.OutLineSummaryRight"/>
        /// <paramref name="allLevels">If true, all children will be expanded and shown. If false, only the children of the referenced columns will be expanded.</paramref>
        /// </summary>
        void ExpandChildren(bool allLevels = true);
        /// <summary>
        /// Expands the columns to the <see cref="OutlineLevel"/> supplied. 
        /// </summary>
        /// <param name="level">Expand all columns with a <see cref="OutlineLevel"/> Equal or Greater than this number.</param>
        /// <param name="collapseChildren">Collapse all children with a greater <see cref="OutlineLevel"/> than <paramref name="level"/></param>
        void SetVisibleOutlineLevel(int level, bool collapseChildren = true);
    }
    /// <summary>
    /// Represents a range of columns
    /// </summary>
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
        /// <summary>
        /// If the column is collapsed in outline mode
        /// </summary>
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
        /// <summary>
        /// Groups the columns using an outline. Adds one to <see cref="OutlineLevel" /> for each column if the outline level is less than 8.
        /// </summary>
        public void Group()
        {
            SetValue(new Action<ExcelColumn, int>((x, v) => { if(x.OutlineLevel<8) x.OutlineLevel += v; }), 1);
        }
        /// <summary>
        /// Ungroups the columns from the outline. 
        /// Subtracts one from <see cref="OutlineLevel" /> for each column if the outline level is larger that zero. 
        /// </summary>
        public void UnGroup()
        {
            SetValue(new Action<ExcelColumn, int>((x, v) => { if (x.OutlineLevel >= 0) x.OutlineLevel += v; }), -1);
        }

        /// <summary>
        /// Collapses and hides the column's children. Children are columns immegetaly to the right or left of the column depending on the <see cref="ExcelWorksheet.OutLineSummaryRight"/>
        /// <paramref name="allLevels">If true, all children will be collapsed and hidden. If false, only the children of the referenced columns are collapsed.</paramref>
        /// </summary>
        public void CollapseChildren(bool allLevels = true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryRight)
            {
                for (int c = GetLastCol(); c >= _fromCol; c--)
                {
                    c = helper.CollapseColumn(c, allLevels ? -1 : -2, true, true, -1);
                }
            }
            else
            {
                for (int c = _fromCol; c <= GetLastCol(); c++)
                {
                    c = helper.CollapseColumn(c, allLevels ? -1 : -2, true, true, 1);
                }
            }            
        }
        /// <summary>
        /// Expands and shows the column's children. Children are columns immegetaly to the right or left of the column depending on the <see cref="ExcelWorksheet.OutLineSummaryRight"/>
        /// <paramref name="allLevels">If true, all children will be expanded and shown. If false, only the children of the referenced columns will be expanded.</paramref>
        /// </summary>
        public void ExpandChildren(bool allLevels = true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryRight)
            {
                for (int c = GetLastCol(); c >= _fromCol; c--)
                {
                    c = helper.CollapseColumn(c, allLevels ? -1 : -2, false, true, -1);
                }
            }
            else
            {
                for (int c = _fromCol; c <= GetLastCol(); c++)
                {
                    c = helper.CollapseColumn(c, allLevels ? -1 : -2, false, true, 1);
                }
            }
        }
        /// <summary>
        /// Expands the rows to the <see cref="OutlineLevel"/> supplied. 
        /// </summary>
        /// <param name="level">Expand all rows with a <see cref="OutlineLevel"/> Equal or Greater than this number.</param>
        /// <param name="collapseChildren">Collapse all children with a greater <see cref="OutlineLevel"/> than <paramref name="level"/></param>
        public void SetVisibleOutlineLevel(int level, bool collapseChildren = true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryRight)
            {
                for (int c = GetLastCol(); c >= _fromCol; c--)
                {
                    c = helper.CollapseColumn(c, level, true, collapseChildren, -1);
                }
            }
            else
            {
                for (int c = _fromCol; c <= GetLastCol(); c++)
                {
                    c = helper.CollapseColumn(c, level, true, collapseChildren, 1);
                }
            }
        }

        /// <summary>
        /// Outline level. Zero if no outline. Can not be negative.
        /// </summary>
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
        /// <summary>
        /// True if the column should show phonetic
        /// </summary>
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
        /// <summary>
        /// Indicates that the column should resize when numbers are entered into the column to fit the size of the text.
        /// This only applies to columns where the size has not been set.
        /// </summary>
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

        /// <summary>
        /// If the column is hidden.
        /// </summary>
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
        /// <summary>
        /// Row width of the column.
        /// </summary>
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
        internal double VisualWidth
        {
            get
            {
                return GetValue(new Func<ExcelColumn, double>(x => x.VisualWidth), _worksheet.DefaultColWidth);
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
        /// <summary>
        /// The current range when enumerating
        /// </summary>
        public ExcelRangeColumn Current
        {
            get
            {
                return new ExcelRangeColumn(_worksheet, enumCol, enumCol);
            }
        }
        /// <summary>
        /// The current range when enumerating
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return new ExcelRangeColumn(_worksheet, enumCol, enumCol);
            }
        }
        #endregion

        /// <summary>
        /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
        /// The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
        /// </summary>
        /// <remarks>
        /// Cells containing formulas must be calculated before autofit is called.
        /// Wrapped and merged cells are also ignored.
        /// </remarks>
        public void AutoFit()
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns();
        }

        /// <summary>
        /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
        /// </summary>
        /// <remarks>
        /// This method will not work if you run in an environment that does not support GDI.
        /// Cells containing formulas are ignored if no calculation is made.
        /// Wrapped and merged cells are also ignored.
        /// </remarks>
        /// <param name="MinimumWidth">Minimum column width</param>
        public void AutoFit(double MinimumWidth)
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth);
        }

        /// <summary>
        /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
        /// </summary>
        /// <remarks>
        /// This method will not work if you run in an environment that does not support GDI.
        /// Cells containing formulas are ignored if no calculation is made.
        /// Wrapped and merged cells are also ignored.
        /// </remarks>        
        /// <param name="MinimumWidth">Minimum column width</param>
        /// <param name="MaximumWidth">Maximum column width</param>
        public void AutoFit(double MinimumWidth, double MaximumWidth)
        {
            _worksheet.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth, MaximumWidth);
        }
        private ExcelColumn GetColumn(int col, bool ignoreFromCol = true)
        {
            var currentCol = _worksheet.GetValueInner(0, col) as ExcelColumn;
            if (currentCol == null)
            {
                int r = 0, c = col;
                if (_worksheet._values.PrevCell(ref r, ref c))
                {
                    if (c > 0)
                    {
                        ExcelColumn prevCol = _worksheet.GetValueInner(0, c) as ExcelColumn;
                        if (prevCol.ColumnMax >= _fromCol || ignoreFromCol)
                        {
                            return prevCol;
                        }
                    }
                }
            }
            return currentCol;
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
            if (currentCol == null)
            {
                int cPrev = _fromCol;
                if (_worksheet._values.PrevCell(ref r, ref cPrev))
                {
                    var pc = _worksheet.GetValueInner(0, cPrev) as ExcelColumn;
                    if (cPrev > 0)
                    {
                        ExcelColumn prevCol = _worksheet.GetValueInner(0, cPrev) as ExcelColumn;
                        if (prevCol.ColumnMax >= _fromCol)
                        {
                            currentCol = prevCol;
                        }
                    }
                }
            }

            while (c <= _toCol)
            {
                if (currentCol == null)
                {
                    currentCol = _worksheet.Column(c);
                }
                else
                {
                    if (c < _fromCol || c != currentCol.ColumnMin)
                    {
                        currentCol = _worksheet.Column(c);
                    }
                }

                if (currentCol.ColumnMax > _toCol)
                {
                    AdjustColumnMaxAndCopy(currentCol, _toCol);
                }
                else if(currentCol.ColumnMax < _toCol)
                {
                    if (_worksheet._values.NextCell(ref r, ref c))
                    {
                        if (r == 0 && c <= _toCol)
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

        private void AdjustColumnMaxAndCopy(ExcelColumn currentCol, int newColMax)
        {
            if (newColMax < currentCol.ColumnMax)
            {
                int maxCol = currentCol.ColumnMax;
                currentCol.ColumnMax = newColMax;
                ExcelColumn copy = _worksheet.CopyColumn(currentCol, newColMax + 1, maxCol);
            }
        }

        /// <summary>
        /// Reference to the cell range of the column(s)
        /// </summary>
        public ExcelRangeBase Range
        {
            get
            {
                return new ExcelRangeBase(_worksheet, ExcelAddressBase.GetAddress(1, _fromCol, ExcelPackage.MaxRows, _toCol));
            }
        }
        /// <summary>
        /// Gets the enumerator
        /// </summary>

        public IEnumerator<ExcelRangeColumn> GetEnumerator()
        {
            return this;
        }
        /// <summary>
        /// Gets the enumerator
        /// </summary>

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }

        /// <summary>
        /// Iterate to the next row
        /// </summary>
        /// <returns>False if no more row exists</returns>
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
                    enumColPos = _cs.GetColumnPosition(enumCol);
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
        /// <summary>
        /// Reset the enumerator
        /// </summary>
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
                    if(enumColPos > 0 && _cs.GetPrevCell(ref r, ref c, 0, enumColPos - 1, _toCol))
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
                        enumCol = _cs._columnIndex[enumColPos].Index;
                        _currentCol = _cs.GetValue(0, enumCol)._value as ExcelColumn;
                    }

                }
                else
                {
                    _currentCol = _cs.GetValue(0, _fromCol)._value as ExcelColumn; 
                }
            }
        }
        private int GetLastCol()
        {
            int maxCol;
            if
                (_worksheet.Dimension == null)
            {
                maxCol = _worksheet._values.GetLastColumn();
            }
            else
            {
                maxCol = Math.Max(_worksheet.Dimension.End.Row, _worksheet._values.GetLastRow(0));
            }
            return _toCol > maxCol + 1 ? maxCol + 1 : _toCol;   // +1 if the last column has outline level 1 then +1 is outline level 0.
        }

        /// <summary>
        /// Disposes this object
        /// </summary>
        public void Dispose()
        {
            
        }
    }
}
