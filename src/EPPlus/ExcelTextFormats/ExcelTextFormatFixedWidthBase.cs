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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
        /// Padding character for Text Can be set to null to skip trimming padding characters when reading
        /// </summary>
        public char PaddingCharacter { get; set; } = ' ';

        /// <summary>
        /// Padding character for numbers.
        /// </summary>
        public char? PaddingCharacterNumeric { get; set; } = null;

        /// <summary>
        /// Set if we should read fixed width files from column widths or positions. Default is widths
        /// </summary>
        public FixedWidthReadType ReadType { get; set; } = FixedWidthReadType.Length;

        int _lineLength;

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

        private void CreateColumnFormatList(int size)
        {
            for (int i =0; i<size; i++)
            {
                ColumnFormat.Add(new ExcelTextFormatColumn());
            }
        }

        /// <summary>
        /// Clear the collection of column formats.
        /// </summary>
        public void ClearColumnFormats()
        {
            _lineLength = 0;
            ColumnFormat.Clear();
        }

        /// <summary>
        /// Adds the column read by column length or position and data.
        /// </summary>
        /// <param name="readType">Specify if lenght or position is provided</param>
        /// <param name="lineLength">The length of a line in the file. If readType is position, a value greater than  0 or negative will read until EOL. If readType is lengths you can ignore this argument</param>
        /// <param name="firstPosition">The start position of the</param>
        /// <param name="columns"></param>
        public void SetColumns(FixedWidthReadType readType, int lineLength = 0, int firstPosition = 0, params int[] columns)
        {
            if (readType == FixedWidthReadType.Length)
            {
                SetColumnLengths(columns);
            }
            else if (readType == FixedWidthReadType.Positions)
            {
                SetColumnPositions(lineLength, firstPosition, columns);
            }
            ReadType = readType;
        }

        /// <summary>
        /// Set the column length properties of fixed width text.
        /// </summary>
        /// <param name="lengths"></param>
        public void SetColumnLengths(params int[] lengths)
        {
            if (ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(lengths.Length);
            }
            for (int i = 0; i < lengths.Length; i++)
            {
                ColumnFormat[i].Length = lengths[i];
                if (ReadType == FixedWidthReadType.Length)
                {
                    _lineLength += lengths[i];
                }
            }
        }

        /// <summary>
        /// Set the column start positions of fixed width text.
        /// </summary>
        /// <param name="lineLength">The Length of a line. Set to 0 or negative to read until end of line</param>
        /// <param name="firstPosition">The starting position of the first column</param>
        /// <param name="positions">Starting positions for the other columns in order from second column</param>
        public void SetColumnPositions(int lineLength, int firstPosition, params int[] positions)
        {
            if (ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(positions.Length + 1);
            }
            ColumnFormat[0].Position = firstPosition;
            for (int i = 0; i < positions.Length; i++)
            {
                if (positions[i] < ColumnFormat[i].Position)
                {
                    throw new ArgumentException("Positions value at " + i + " was lower that previous value " + ColumnFormat[i].Position);
                }
                ColumnFormat[i+1].Position = positions[i];
            }
            if(lineLength > 0 && lineLength > ColumnFormat[ColumnFormat.Count - 1].Position)
            {
                var lastPosLen = lineLength - ColumnFormat[ColumnFormat.Count - 1].Position;
                ColumnFormat[ColumnFormat.Count - 1].Length = lastPosLen;
                _lineLength = lineLength;
                return;
            }
            else
            {
                if(lineLength <= 0)
                {
                    _lineLength = ColumnFormat[ColumnFormat.Count - 1].Position;
                    return;
                }
                else
                {
                    throw new ArgumentException("lineLength cannot be smaller than last supplied position");
                }
            }
        }

        /// <summary>
        /// Set the data types for each column.
        /// </summary>
        /// <param name="dataTypes"></param>
        public void SetColumnDataTypes(params eDataTypes[] dataTypes)
        {
            int i = 0;
            if (ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(dataTypes.Length);
            }
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
        public void SetColumnPaddingAlignmentType(params PaddingAlignmentType[] paddingTypes)
        {
            int i = 0;
            if(ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(paddingTypes.Length);
            }
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
            if (ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(UseColumns.Length);
            }
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
            if (ColumnFormat.Count <= 0)
            {
                CreateColumnFormatList(Names.Length);
            }
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
