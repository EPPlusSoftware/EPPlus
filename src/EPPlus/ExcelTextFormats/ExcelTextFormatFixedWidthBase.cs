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
        Length,
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

        /// <summary>
        /// The collection of column formats.
        /// </summary>
        public List<ExcelTextFormatColumn> ColumnFormat { get; set; } = new List<ExcelTextFormatColumn>();

        /// <summary>
        /// Force writing to file, this will only write the n first found characters, where n is column width
        /// </summary>
        public bool ForceWrite { get; set; } = false;

        /// <summary>
        /// Force Reading content. Setting this will force reading line a line that is not following column spec.
        /// </summary>
        public bool ForceRead { get; set; } = false;
        /// <summary>
        /// 
        /// </summary>
        public char PaddingCharacter { get; set; } = ' ';

        /// <summary>
        /// Set if we should read fixed width files from column widths or positions. Default is widths
        /// </summary>
        public FixedWidthReadType ReadStartPosition { get; set; } = FixedWidthReadType.Length;

        int _lineLength;
        int _lastPosition;

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
        /// The position of the last column.
        /// </summary>
        public int LastPosition
        {
            get
            {
                return _lastPosition;
            }
            set
            {
                _lastPosition = value;
            }
        }

        /// <summary>
        /// Adds the column read by column length or position and data.
        /// </summary>
        /// <param name="readType"></param>
        /// <param name="columns"></param>
        public void SetColumns(FixedWidthReadType readType, params int[] columns)
        {
            ColumnFormat.Clear();
            if (readType == FixedWidthReadType.Length)
            {
                foreach (int column in columns)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Length = column });
                    _lineLength += column;
                    _lastPosition = _lineLength;
                }
            }
            else if (readType == FixedWidthReadType.Positions)
            {
                foreach (int column in columns)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Position = column });
                    _lineLength = column;
                    _lastPosition = column;
                }
            }
            ReadStartPosition = readType;
        }

        /// <summary>
        /// Set the column length properties of fixed width text.
        /// </summary>
        /// <param name="lengths"></param>
        public void SetColumnLengths(params int[] lengths)
        {
            int i = 0;
            foreach (int length in lengths)
            {
                if (ColumnFormat.Count <= i)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Length = length });
                    _lineLength += length;
                    _lastPosition = _lineLength;
                }
                ColumnFormat[i].Length = length;
                i++;
            }
        }

        /// <summary>
        /// Set the column start positions of fixed width text.
        /// </summary>
        /// <param name="positions"></param>
        public void SetColumnPositions(params int[] positions)
        {
            int i = 0;
            foreach (int position in positions)
            {
                if (ColumnFormat.Count <= i)
                {
                    ColumnFormat.Add(new ExcelTextFormatColumn() { Position = position });
                    _lineLength = position;
                    _lastPosition = position;
                }        
                ColumnFormat[i].Position = position;
                i++;
            }
        }

        /// <summary>
        /// Set the data types for each column.
        /// </summary>
        /// <param name="dataTypes"></param>
        public void SetColumnDataTypes(params eDataTypes[] dataTypes)
        {
            int i = 0;
            foreach (eDataTypes dataType in dataTypes)
            {
                if (ColumnFormat.Count <= i)
                {
                    return;
                }
                ColumnFormat[i].DataType = dataType;
                i++;
            }
        }

        /// <summary>
        /// Set the padding type for each column. 
        /// </summary>
        /// <param name="paddingTypes"></param>
        public void SetColumnPaddingType(params int[] paddingTypes)
        {
            int i = 0;
            foreach (PaddingAlignmentType paddingType in paddingTypes)
            {
                if (ColumnFormat.Count <= i)
                {
                    return;
                }
                ColumnFormat[i].PaddingType = paddingType;
                i++;
            }
        }

        /// <summary>
        /// Set flag for each column to be used. 
        /// </summary>
        /// <param name="UseColumns"></param>
        public void SetUseColumns(params bool[] UseColumns)
        {
            int i = 0;
            foreach (bool UseColumn in UseColumns)
            {
                if (ColumnFormat.Count <= i)
                {
                    return;
                }
                ColumnFormat[i].UseColumn = UseColumn;
                i++;
            }
        }

        /// <summary>
        /// Set flag for each column to be used. 
        /// </summary>
        /// <param name="Names"></param>
        public void SetColumnsNames(params string[] Names)
        {
            int i = 0;
            foreach (string name in Names)
            {
                if (ColumnFormat.Count <= i)
                {
                    return;
                }
                ColumnFormat[i].Name = name;
                i++;
            }
        }
    }
}
