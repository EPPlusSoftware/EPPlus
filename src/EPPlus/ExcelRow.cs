/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
	internal class RowInternal
    {
        internal double Height=-1;
        internal bool Hidden;
        internal bool Collapsed;        
        internal short OutlineLevel;
        internal bool PageBreak;
        internal bool Phonetic;
        internal bool CustomHeight;
        internal int MergeID;
        internal RowInternal Clone()
        {
            return new RowInternal()
            {
                Height=Height,
                Hidden=Hidden,
                Collapsed=Collapsed,
                OutlineLevel=OutlineLevel,
                PageBreak=PageBreak,
                Phonetic=Phonetic,
                CustomHeight=CustomHeight,
                MergeID=MergeID
            };
        }
    }
    /// <summary>
	/// Represents an individual row in the spreadsheet.
	/// </summary>
	public class ExcelRow : IRangeID
	{
		private ExcelWorksheet _worksheet;
		private XmlElement _rowElement = null;
        /// <summary>
        /// Internal RowID.
        /// </summary>
        [Obsolete]
        public ulong RowID 
        {
            get
            {
                return GetRowID(_worksheet.SheetId, Row);
            }
        }
		#region ExcelRow Constructor
		/// <summary>
		/// Creates a new instance of the ExcelRow class. 
		/// For internal use only!
		/// </summary>
		/// <param name="Worksheet">The parent worksheet</param>
		/// <param name="row">The row number</param>
		internal ExcelRow(ExcelWorksheet Worksheet, int row)
		{
			_worksheet = Worksheet;
            Row = row;
		}
		#endregion

		/// <summary>
		/// Provides access to the node representing the row.
		/// </summary>
		internal XmlNode Node { get { return (_rowElement); } }

		#region ExcelRow Hidden
        /// <summary>
		/// Allows the row to be hidden in the worksheet
		/// </summary>
		public bool Hidden
        {
            get
            {
                var r=(RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return false;
                }
                else
                {
                    return r.Hidden;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.Hidden=value;
            }
        }        
		#endregion

		#region ExcelRow Height
        /// <summary>
		/// Sets the height of the row
		/// </summary>
		public double Height
        {
			get
			{
                var r = (RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null || r.Height<0)
                {
                    return _worksheet.DefaultRowHeight;
                }
                else
                {
                    return r.Height;
                }
            }
            set
            {
                var r = GetRowInternal();
                if (_worksheet._package.DoAdjustDrawings)
                {
                    var pos = _worksheet.Drawings.GetDrawingHeight();   //Fixes issue 14846
                    _worksheet.RowHeightCache.Remove(Row - 1);
                    r.Height = value;
                    _worksheet.Drawings.AdjustHeight(pos);
                }
                else
                {
                    r.Height = value;
                }
                
                if (r.Hidden && value != 0)
                {
                    Hidden = false;
                }
                r.CustomHeight = true;
            }
        }
        /// <summary>
        /// Set to true if You don't want the row to Autosize
        /// </summary>
        public bool CustomHeight 
        {
            get
            {
                var r = (RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return false;
                }
                else
                {
                    return r.CustomHeight;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.CustomHeight = value;
            }
        }
		#endregion

        internal string _styleName = "";
        /// <summary>
        /// Sets the style for the entire column using a style name.
        /// </summary>
        public string StyleName
        {
            get
            {
                return _styleName;
            }
            set
            {
                StyleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
        }
        /// <summary>
        /// Sets the style for the entire row using the style ID.  
        /// </summary>
        public int StyleID
        {
            get
            {
                return _worksheet.GetStyleInner(Row, 0);
            }
            set
            {
                _worksheet.SetStyleInner(Row, 0, value);
            }
        }

        /// <summary>
        /// Rownumber
        /// </summary>
        public int Row
        {
            get;
            set;
        }
        /// <summary>
        /// If outline level is set this tells that the row is collapsed
        /// </summary>
        public bool Collapsed
        {
            get
            {
                var r=(RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return false;
                }
                else
                {
                    return r.Collapsed;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.Collapsed = value;
            }
        }
        /// <summary>
        /// Outline level.
        /// </summary>
        public int OutlineLevel
        {
            get
            {
                var r=(RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return 0;
                }
                else
                {
                    return r.OutlineLevel;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.OutlineLevel=(short)value;
            }
        }

        private RowInternal GetRowInternal()
        {
            return GetRowInternal(_worksheet, Row);
        }        
        internal static RowInternal GetRowInternal(ExcelWorksheet ws, int row)
        {
            var r = (RowInternal)ws.GetValueInner(row, 0);
            if (r == null)
            {
                r = new RowInternal();
                ws.SetValueInner(row, 0, r);
            }
            return r;
        }
        /// <summary>
        /// Show phonetic Information
        /// </summary>
        public bool Phonetic 
        {
            get
            {
                var r = (RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return false;
                }
                else
                {
                    return r.Phonetic;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.Phonetic = value;
            }
        }
        /// <summary>
        /// The Style applied to the whole row. Only effekt cells with no individual style set. 
        /// Use the <see cref="ExcelWorksheet.Cells"/> Style property if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID,_worksheet.PositionId ,Row.ToString(CultureInfo.InvariantCulture) + ":" + Row.ToString(CultureInfo.InvariantCulture));                
            }
        }
        /// <summary>
        /// Adds a manual page break after the row.
        /// </summary>
        public bool PageBreak
        {
            get
            {
                var r = (RowInternal)_worksheet.GetValueInner(Row, 0);
                if (r == null)
                {
                    return false;
                }
                else
                {
                    return r.PageBreak;
                }
            }
            set
            {
                var r = GetRowInternal();
                r.PageBreak = value;
            }
        }
        /// <summary>
        /// Merge all cells in the row
        /// </summary>
        public bool Merged
        {
            get
            {
                return _worksheet.MergedCells[Row, 0] != null;
            }
            set
            {
                _worksheet.MergedCells.Add(new ExcelAddressBase(Row, 1, Row, ExcelPackage.MaxColumns), true);
            }
        }
        internal static ulong GetRowID(int sheetID, int row)
        {
            return ((ulong)sheetID) + (((ulong)row) << 29);

        }
        
        #region IRangeID Members

        [Obsolete]
        ulong IRangeID.RangeID
        {
            get
            {
                return RowID; 
            }
            set
            {
                Row = ((int)(value >> 29));
            }
        }

        #endregion
        /// <summary>
        /// Copies the current row to a new worksheet
        /// </summary>
        /// <param name="added">The worksheet where the copy will be created</param>
        internal void Clone(ExcelWorksheet added)
        {
            var rowSource = _worksheet.GetValue(Row, 0) as RowInternal;
            if(rowSource != null)
            {
                added.SetValueInner(Row, 0, rowSource.Clone());
            }
        }
    }
}
