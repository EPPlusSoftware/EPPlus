using OfficeOpenXml.Style;
using System;

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
    public class ExcelRangeColumn : IExcelColumn
    {
        ExcelWorksheet _ws;
        internal int _fromCol, _toCol;
        internal ExcelRangeColumn(ExcelWorksheet ws, int fromCol, int toCol)
        {
            _ws = ws;
            _fromCol = fromCol;
            _toCol = toCol;            
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
                return GetValue(new Func<ExcelColumn, double>(x => x.Width), _ws.DefaultColWidth);
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
        #endregion

        public void AutoFit()
        {
            _ws.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns();
        }

        public void AutoFit(double MinimumWidth)
        {
            _ws.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth);
        }

        public void AutoFit(double MinimumWidth, double MaximumWidth)
        {
            _ws.Cells[1, _fromCol, ExcelPackage.MaxRows, _toCol].AutoFitColumns(MinimumWidth, MaximumWidth);
        }
        private TOut GetValue<TOut>(Func<ExcelColumn, TOut> SetValue, TOut defaultValue)
        {
            var currentCol = _ws.GetValueInner(0, _fromCol) as ExcelColumn;
            if (currentCol == null)
            {
                return defaultValue;
            }
            else
            {
                return SetValue(currentCol);
            }
        }

        private void SetValue<T>(Action<ExcelColumn,T> SetValue, T value)
        {
            var c = _fromCol;
            int r = 0;
            ExcelColumn currentCol = _ws.GetValueInner(0, c) as ExcelColumn;
            while (c <= _toCol)
            {
                if (currentCol == null)
                {
                    currentCol = _ws.Column(c);
                }
                else
                {
                    if (c < _fromCol)
                    {
                        currentCol = _ws.Column(c);
                    }
                }

                if (currentCol.ColumnMax >= _toCol)
                {
                    currentCol.ColumnMax = _toCol;
                }
                else
                {
                    if (_ws._values.NextCell(ref r, ref c))
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
    }
}
