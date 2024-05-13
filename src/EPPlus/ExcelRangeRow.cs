using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml
{
    /// <summary>
    /// A row in a worksheet
    /// </summary>
    interface IExcelRow
    {
        /// <summary>
        /// If the row is collapsed in outline mode
        /// </summary>
        bool Collapsed { get; set; }
        /// <summary>
        /// Outline level. Zero if no outline
        /// </summary>
        int OutlineLevel { get; set; }
        /// <summary>
        /// True if the row should show phonetic
        /// </summary>
        bool Phonetic { get; set; }
        /// <summary>
        /// If the row is hidden.
        /// </summary>
        bool Hidden
        {
            get;
            set;
        }
        /// <summary>
        /// Row height in points if specified manually.
        /// <seealso cref="CustomHeight"/>
        /// </summary>
        double Height
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
        /// True if height is set manually
        /// </summary>
        bool CustomHeight
        {
            get;
            set;
        }
        /// <summary>
        /// Groups the rows using an outline. 
        /// Adds one to <see cref="OutlineLevel" /> for each row if the outline level is less than 8.
        /// </summary>
        void Group();
        /// <summary>
        /// Ungroups the rows from the outline. 
        /// Subtracts one from <see cref="OutlineLevel" /> for each row if the outline level is larger that zero. 
        /// </summary>
        void Ungroup();
        /// <summary>
        /// Collapses and hides the rows's children. Children are rows immegetaly below or top of the row depending on the <see cref="ExcelWorksheet.OutLineSummaryBelow"/>
        /// <paramref name="allLevels">If true, all children will be collapsed and hidden. If false, only the children of the referenced rows are collapsed.</paramref>
        /// </summary>
        void CollapseChildren(bool allLevels = true);
        /// <summary>
        /// Expands and shows the rows's children. Children are columns immegetaly below or top of the row depending on the <see cref="ExcelWorksheet.OutLineSummaryBelow"/>
        /// <paramref name="allLevels">If true, all children will be expanded and shown. If false, only the children of the referenced columns will be expanded.</paramref>
        /// </summary>
        void ExpandChildren(bool allLevels = true);
        /// <summary>
        /// Expands the rows to the <see cref="OutlineLevel"/> supplied. 
        /// </summary>
        /// <param name="level">Expands all rows with a <see cref="OutlineLevel"/> Equal or Greater than this number.</param>
        /// <param name="collapseChildren">Collapses all children with a greater <see cref="OutlineLevel"/> than <paramref name="level"/></param>
        void SetVisibleOutlineLevel(int level, bool collapseChildren = true);
    }
    /// <summary>
    /// Represents a range of rows
    /// </summary>
    public class ExcelRangeRow : IExcelRow, IEnumerable<ExcelRangeRow>, IEnumerator<ExcelRangeRow>
    {
        ExcelWorksheet _worksheet;
        internal int _fromRow, _toRow;
        internal ExcelRangeRow(ExcelWorksheet worksheet, int fromRow, int toRow)
        {
            _worksheet = worksheet;
            _fromRow = fromRow;
            _toRow = toRow;
        }
        /// <summary>
        /// The first row in the collection
        /// </summary>
        public int StartRow
        { 
            get
            {
                return _fromRow;
            }
        }
        /// <summary>
        /// The last row in the collection
        /// </summary>
        public int EndRow
        {
            get
            {
                return _toRow;
            }
        }
        /// <summary>
        /// If the row is collapsed in outline mode
        /// </summary>
        public bool Collapsed
        {
            get
            {
                return GetValue(new Func<RowInternal, bool>(x => x.Collapsed), false);
            }
            set
            {
                SetValue(new Action<RowInternal, bool>((x, v) => { x.Collapsed = v; }), value);
            }
        }
        /// <summary>
        /// Outline level. Zero if no outline
        /// </summary>
        public int OutlineLevel
        {
            get
            {
                return GetValue(new Func<RowInternal, int>(x => x.OutlineLevel), 0);
            }
            set
            {
                SetValue(new Action<RowInternal, int>((x, v) => { x.OutlineLevel = (short)v; }), value);
            }
        }

        /// <summary>
        /// True if the row should show phonetic
        /// </summary>
        public bool Phonetic
        {
            get
            {
                return GetValue(new Func<RowInternal, bool>(x => x.Phonetic), false);
            }
            set
            {
                SetValue(new Action<RowInternal, bool>((x, v) => { x.Phonetic = v; }), value);
            }
        }
        /// <summary>
        /// If the row is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                return GetValue(new Func<RowInternal, bool>(x => x.Hidden), false);
            }
            set
            {
                SetValue(new Action<RowInternal, bool>((x, v) => { x.Hidden = v; }), value);
            }
        }

        /// <summary>
        /// Row height in points. Setting this property will also set <see cref="CustomHeight"/> to true.
        /// </summary>
        public double Height
        {
            get
            {
                return GetValue(new Func<RowInternal, double>(x => x.Height), _worksheet.DefaultRowHeight);
            }
            set
            {
                SetValue(new Action<RowInternal, double>((x, v) => 
                { 
                    x.Height = v;
                    x.CustomHeight = true; 
                }), value);
            }
        }
        /// <summary>
        /// True if the row <see cref="Height" /> has been manually set.
        /// </summary>
        public bool CustomHeight
        {
            get
            {
                return GetValue(new Func<RowInternal, bool>(x => x.CustomHeight), false);
            }
            set
            {
                SetValue(new Action<RowInternal, bool>((x, v) => { x.CustomHeight = v; }), value);
            }
        }

        /// <summary>
        /// Adds a manual page break after the column.
        /// </summary>
        public bool PageBreak
        {
            get
            {
                return GetValue(new Func<RowInternal, bool>(x => x.PageBreak), false);
            }
            set
            {
                SetValue(new Action<RowInternal, bool>((x, v) => { x.PageBreak = v; }), value);
            }
        }
        #region ExcelRow Style
        /// <summary>
        /// The Style applied to the whole row(s). Only effects cells with no individual style set. 
        /// Use the Range object if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionId, _fromRow.ToString(CultureInfo.InvariantCulture) + ":" + _toRow.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal string _styleName = "";
        /// <summary>
		/// Sets the style for the entire row using a style name.
		/// </summary>
		public string StyleName
        {

            get
            {
                var xfId = _worksheet.Workbook.Styles.CellXfs[StyleID].XfId;
                if (xfId >= 0 && xfId < _worksheet.Workbook.Styles.CellStyleXfs.Count)
                {
                    var ns = _worksheet.Workbook.Styles.NamedStyles.Where(x => x.StyleXfId == xfId).FirstOrDefault();
                    if (ns != null)
                    {
                        return ns.Name;
                    }
                }
                return "";
            }
            set
            {
                StyleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
            }
        }
        /// <summary>
        /// Sets the style for the entire column using the style ID.           
        /// </summary>
        public int StyleID
        {
            get
            {
                return _worksheet.GetStyleInner(_fromRow, 0);
            }
            set
            {
                for (int r = _fromRow; r <= _toRow; r++)
                {
                    _worksheet.SetStyleInner(r, 0, value);
                }
            }
        }
        /// <summary>
        /// Reference to the cell range of the row(s)
        /// </summary>
        public ExcelRangeBase Range
        {
            get
            {
                return new ExcelRangeBase(_worksheet, ExcelAddressBase.GetAddress(_fromRow, 1, _toRow, ExcelPackage.MaxColumns));
            }
        }
        /// <summary>
        /// The current row object in the iteration
        /// </summary>
        public ExcelRangeRow Current
        {
            get
            {
                return new ExcelRangeRow(_worksheet, enumRow, enumRow);
            }
        }

        /// <summary>
        /// The current row object in the iteration
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return new ExcelRangeRow(_worksheet, enumRow, enumRow);
            }
        }


        #endregion

        private TOut GetValue<TOut>(Func<RowInternal, TOut> getValue, TOut defaultValue)
        {
            var currentRow = _worksheet.GetValueInner(_fromRow, 0) as RowInternal;
            if (currentRow == null)
            {
                return defaultValue;
            }
            else
            {
                return getValue(currentRow);
            }
        }

        private void SetValue<T>(Action<RowInternal,T> SetValue, T value)
        {
            for(int r=_fromRow;r<=_toRow;r++)
            {
                var row = _worksheet.GetValueInner(r, 0) as RowInternal;
                if(row==null)
                {
                    row = new RowInternal();
                    _worksheet.SetValueInner(r, 0, row);
                }
                SetValue(row, value);
            }
        }

        /// <summary>
        /// Gets the enumerator
        /// </summary>
        public IEnumerator<ExcelRangeRow> GetEnumerator()
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

        CellStoreValue _cs;
        int enumRow = -1;
        int enumCol = -1;
        int minCol=-1;
        /// <summary>
        /// Iterate to the next row
        /// </summary>
        /// <returns>False if no more row exists</returns>
        public bool MoveNext()
        {
            if (minCol < 0)
            {
                if (_cs == null) Reset();
                if (minCol < 0) return false;
            }
            enumCol = -1;
            return _cs.NextCell(ref enumRow, ref enumCol, enumRow, minCol, _toRow,0);
        }

        /// <summary>
        /// Reset the enumerator
        /// </summary>
        public void Reset()
        {
            _cs = _worksheet._values;
            enumRow = _fromRow - 1;
            minCol = 0;
        }
        /// <summary>
        /// Disposes this object
        /// </summary>
        public void Dispose()
        {
        }
        /// <summary>
        /// Groups the rows using an outline. 
        /// Adds one to <see cref="OutlineLevel" /> for each row if the outline level is less than 8.
        /// </summary>
        public void Group()
        {
            SetValue(new Action<RowInternal, int>((x, v) => { if (x.OutlineLevel < 8) x.OutlineLevel += (short)v; }), 1);
        }
        /// <summary>
        /// Ungroups the rows from the outline. 
        /// Subtracts one from <see cref="OutlineLevel" /> for each row if the outline level is larger that zero. 
        /// </summary>
        public void Ungroup()
        {
            SetValue(new Action<RowInternal, int>((x, v) => { if (x.OutlineLevel >= 0) x.OutlineLevel += (short)v; }), -1);
        }
        /// <summary>
        /// Collapses and hides the rows's children. Children are rows immegetaly below or top of the row depending on the <see cref="ExcelWorksheet.OutLineSummaryBelow"/>
        /// <paramref name="allLevels">If true, all children will be collapsed and hidden. If false, only the children of the referenced rows are collapsed.</paramref>
        /// </summary>
        public void CollapseChildren(bool allLevels = true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryBelow)
            {
                for (int c = GetToRow(); c >= _fromRow; c--)
                {
                    c = helper.CollapseRow(c, allLevels ? -1 : -2, true, true, -1);
                }
            }
            else
            {
                for (int c = _fromRow; c <= GetToRow(); c++)
                {
                    c = helper.CollapseRow(c, allLevels ? -1 : -2, true, true, 1);
                }
            }
        }
        /// <summary>
        /// Expands and shows the rows's children. Children are columns immegetaly below or top of the row depending on the <see cref="ExcelWorksheet.OutLineSummaryBelow"/>
        /// <paramref name="allLevels">If true, all children will be expanded and shown. If false, only the children of the referenced columns will be expanded.</paramref>
        /// </summary>
        public void ExpandChildren(bool allLevels = true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryBelow)
            {
                for (int row = GetToRow(); row >= _fromRow; row--)
                {
                    row = helper.CollapseRow(row, allLevels ? -1 : -2, false, true, -1);
                }
            }
            else
            {
                for (int c = _fromRow; c <= GetToRow(); c++)
                {
                    c = helper.CollapseRow(c, allLevels ? -1 : -2, false, true, 1);
                }
            }
        }
        /// <summary>
        /// Expands the rows to the <see cref="OutlineLevel"/> supplied. 
        /// </summary>
        /// <param name="level">Expand all rows with a <see cref="OutlineLevel"/> Equal or Greater than this number.</param>
        /// <param name="collapseChildren">Collapse all children with a greater <see cref="OutlineLevel"/> than <paramref name="level"/></param>
        public void SetVisibleOutlineLevel(int level, bool collapseChildren=true)
        {
            var helper = new WorksheetOutlineHelper(_worksheet);
            if (_worksheet.OutLineSummaryBelow)
            {
                for (int r = GetToRow(); r >= _fromRow; r--)
                {
                    r = helper.CollapseRow(r, level, true, collapseChildren, -1);
                }
            }
            else
            {
                for (int r = _fromRow; r <= GetToRow(); r++)
                {
                    r = helper.CollapseRow(r, level, true, collapseChildren, 1);
                }
            }
        }
        private int GetToRow()
        {
            int maxRow;
            if 
                (_worksheet.Dimension == null)
            {
                maxRow=_worksheet._values.GetLastRow(0);
            }
            else
            {
                maxRow = Math.Max(_worksheet.Dimension.End.Row, _worksheet._values.GetLastRow(0));
            }
            return _toRow > maxRow + 1 ? maxRow + 1 : _toRow; // +1 if the last row has outline level 1 then +1 is outline level 0.
        }

        private RowInternal GetRow(int row)
        {
            if (row < 1 || row > ExcelPackage.MaxRows) return null;
            return _worksheet.GetValueInner(row, 0) as RowInternal;
        }

    }
}
