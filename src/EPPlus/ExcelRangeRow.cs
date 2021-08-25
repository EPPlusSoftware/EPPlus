using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml
{
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
        /// Phonetic
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
        ///// <summary>
        ///// Merges all cells of the column
        ///// </summary>
        //bool Merged
        //{
        //    get;
        //    set;
        //}
        /// <summary>
        /// Merges all cells of the column
        /// </summary>
        bool CustomHeight
        {
            get;
            set;
        }        
    }
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
        public int StartRow
        { 
            get
            {
                return _fromRow;
            }
        }
        public int EndRow
        {
            get
            {
                return _toRow;
            }
        }
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
        public double Height
        {
            get
            {
                return GetValue(new Func<RowInternal, double>(x => x.Height), _worksheet.DefaultRowHeight);
            }
            set
            {
                SetValue(new Action<RowInternal, double>((x, v) => { x.Height = v; }), value);
            }
        }
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
        ///// <summary>
        ///// Merges all cells of the column
        ///// </summary>
        //public bool Merged
        //{
        //    get
        //    {
        //        return GetValue(new Func<RowInternal, bool>(x => x.), false);
        //    }
        //    set
        //    {
        //        SetValue(new Action<RowInternal, bool>((x, v) => { x.Merged = v; }), value);
        //    }
        //}
        #region ExcelColumn Style
        /// <summary>
        /// The Style applied to the whole column(s). Only effects cells with no individual style set. 
        /// Use Range object if you want to set specific styles.
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
		/// Sets the style for the entire column using a style name.
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
        public ExcelRangeBase Range
        {
            get
            {
                return new ExcelRangeBase(_worksheet, ExcelAddressBase.GetAddress(_fromRow, 1, _toRow, ExcelPackage.MaxColumns));
            }
        }
        public ExcelRangeRow Current
        {
            get
            {
                return new ExcelRangeRow(_worksheet, enumRow, enumRow);
            }
        }

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

        public IEnumerator<ExcelRangeRow> GetEnumerator()
        {
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }

        CellStoreValue _cs;
        int enumRow = -1;
        int enumCol = -1;
        int minCol=-1;
        public bool MoveNext()
        {
            if (minCol < 0)
            {
                if (_cs == null) Reset();
                if (minCol < 0) return false;
            }
            enumCol = -1;
            return _cs.NextCell(ref enumRow, ref enumCol, enumRow, minCol, _toRow,ExcelPackage.MaxColumns);
        }

        public void Reset()
        {
            _cs = _worksheet._values;
            enumRow = _fromRow-1;
            minCol = 0;
        }

        public void Dispose()
        {
        }
    }
}
