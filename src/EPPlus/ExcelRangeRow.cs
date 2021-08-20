using OfficeOpenXml.Style;
using System;
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
    public class ExcelRangeRow : IExcelRow
    {
        ExcelWorksheet _ws;
        internal int _fromRow, _toRow;
        internal ExcelRangeRow(ExcelWorksheet ws, int fromRow, int toRow)
        {
            _ws = ws;
            _fromRow = fromRow;
            _toRow = toRow;            
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
                return GetValue(new Func<RowInternal, double>(x => x.Height), _ws.DefaultRowHeight);
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
                string letter = ExcelCellBase.GetColumnLetter(_fromRow);
                string endLetter = ExcelCellBase.GetColumnLetter(_toRow);
                return _ws.Workbook.Styles.GetStyleObject(StyleID, _ws.PositionId, letter + ":" + endLetter);
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
                var xfId = _ws.Workbook.Styles.CellXfs[StyleID].XfId;
                if(xfId >= 0 && xfId < _ws.Workbook.Styles.CellStyleXfs.Count)
                {
                    var ns = _ws.Workbook.Styles.NamedStyles.Where(x => x.XfId == _ws.Workbook.Styles.CellStyleXfs[xfId].XfId).FirstOrDefault();
                    if(ns!=null)
                    {
                        return ns.Name;
                    }
                }
                return "";
            }
            set
            {
                StyleID = _ws.Workbook.Styles.GetStyleIdFromName(value);
            }
        }
        /// <summary>
        /// Sets the style for the entire column using the style ID.           
        /// </summary>
        public int StyleID
        {
            get
            {
                return _ws.GetStyleInner(_fromRow, 0);
            }
            set
            {
                for (int r = _fromRow; r <= _toRow; r++)
                {
                    _ws.SetStyleInner(r, 0, value);
                }
            }
        }

        #endregion

        private TOut GetValue<TOut>(Func<RowInternal, TOut> getValue, TOut defaultValue)
        {
            var currentRow = _ws.GetValueInner(_fromRow, 0) as RowInternal;
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
                var row = _ws.GetValueInner(r, 0) as RowInternal;
                if(row==null)
                {
                    row = new RowInternal();
                    _ws.SetValueInner(r, 0, row);
                }
                SetValue(row, value);
            }
        }
    }
}
