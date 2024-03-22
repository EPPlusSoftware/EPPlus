/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{

    /// <summary>
    /// 
    /// </summary>
    public enum FixedWidthReadType
    {
        /// <summary>
        /// 
        /// </summary>
        Widths,
        /// <summary>
        /// 
        /// </summary>
        Positions,
    }

    /// <summary>
    /// 
    /// </summary>
    public class ExcelTextFormatFixedWidthBase : ExcelAbstractTextFormat
    {
        int _lineLength;

        /// <summary>
        /// Creates a new instance if ExcelTextFormatFixedWidthBase
        /// </summary>
        public ExcelTextFormatFixedWidthBase() : base()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnsLength"></param>
        public ExcelTextFormatFixedWidthBase(FixedWidthReadType readType, params int[] columns) : base()
        {
            if (readType == FixedWidthReadType.Widths)
            {
                foreach (int column in columns)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Length = column });
                }
            }
            else if(readType == FixedWidthReadType.Positions)
            {
                foreach (int column in columns)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Position = column });
                }
            }
            ReadStartPosition = readType;
        }


        /// <summary>
        /// 
        /// </summary>
        public List<ExcelTextFormatColumn> ColumnFormat { get; set; } = new List<ExcelTextFormatColumn>();

        /// <summary>
        /// The length of the line to read. If set to widths, LineLength is sum of all columnLengths. If set to positions, LineLength is set to the value of the last index of columnLengths
        /// </summary>
        public int LineLength 
        { 
            get 
            { 
                return _lineLength; 
            }
            set
            {
                _lineLength = value;
            }
        }

        /// <summary>
        /// Set if we should read fixed width files from column widths or positions. Default is widths
        /// </summary>
        public FixedWidthReadType ReadStartPosition { get; set; } = FixedWidthReadType.Widths;
    }
}
