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
using OfficeOpenXml.Core;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal enum FixedWidthReadType
    {
        Length,
        Positions,
    }

    /// <summary>
    /// How to handle missmatch with data and column format specifications.
    /// </summary>
    public enum FixedWidthFormatErrorStrategy
    {
        /// <summary>
        /// If data is larger than format specification, write anyway. Might lead to data loss.
        /// </summary>
        Truncate,
        /// <summary>
        /// If data is larger than format specification, throw error
        /// </summary>
        ThrowError,
    }

    /// <summary>
    /// Describes how to import a fixed width text file.
    /// </summary>
    public class ExcelTextFormatFixedWidthBase : ExcelTextFileFormat
    {

        /// <summary>
        /// The collection of column formats.
        /// </summary>
        public EPPlusReadOnlyList<ExcelTextFormatColumn> Columns { get; set; } = new EPPlusReadOnlyList<ExcelTextFormatColumn>();

        /// <summary>
        /// The strategy to use when writing fixed width text files.
        /// </summary>
        public FixedWidthFormatErrorStrategy FormatErrorStrategy { get; set; } = FixedWidthFormatErrorStrategy.ThrowError;

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
        internal FixedWidthReadType ReadType { get; set; } = FixedWidthReadType.Length;

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
                Columns.Add(new ExcelTextFormatColumn());
            }
        }

        /// <summary>
        /// Clear the collection of column formats.
        /// </summary>
        public void ClearColumnFormats()
        {
            _lineLength = 0;
            Columns.Clear();
        }

        /// <summary>
        /// Set the column length properties of fixed width text. For reading to end of line, set last index to 0 or a negative number.
        /// </summary>
        /// <param name="lengths"></param>
        public void SetColumnLengths(params int[] lengths)
        {
            ReadType = FixedWidthReadType.Length;
            if (Columns.Count <= 0)
            {
                CreateColumnFormatList(lengths.Length);
            }
            for (int i = 0; i < lengths.Length; i++)
            {
                if (lengths[i] <= 0 && i != lengths.Length - 1)
                {
                    throw new ArgumentException("Only last column can be 0 or negative.");
                }
                Columns[i]._length = lengths[i];
                _lineLength += lengths[i];
            }
        }

        /// <summary>
        /// Set the column start positions of fixed width text.
        /// </summary>
        /// <param name="lineLength">The Length of a line. Set to 0 or negative to read until end of line.</param>
        /// <param name="positions">The positions where each column starts. This array requires atleast one element.</param>
        public void SetColumnPositions(int lineLength, params int[] positions)
        {
            if(positions.Length <= 0)
            {
                throw new ArgumentException("Requires at least 1 element in positions.");
            }
            for (int i = 0; i < positions.Length-1; i++)
            {
                if (positions[i] >= positions[i+1])
                {
                    throw new ArgumentException("Positions value at " + i + " was lower that previous value " + positions[i]);
                }
            }
            if (Columns.Count <= 0)
            {
                CreateColumnFormatList(positions.Length);
            }
            for (int i = 0; i < positions.Length; i++)
            {
                Columns[i]._position = positions[i];
            }
            ReadType = FixedWidthReadType.Positions;
            if (lineLength > 0 && lineLength > Columns[Columns.Count - 1].Position)
            {
                var lastPosLen = lineLength - Columns[Columns.Count - 1].Position;
                Columns[Columns.Count - 1]._length = lastPosLen;
                _lineLength = lineLength;
                return;
            }
            else
            {
                if(lineLength <= 0)
                {
                    _lineLength = Columns[Columns.Count - 1].Position;
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
            if (Columns.Count <= 0)
            {
                CreateColumnFormatList(dataTypes.Length);
            }
            if(dataTypes.Length > Columns.Count)
            {
                throw new ArgumentException("dataTypes has more elements than Columns");
            }
            foreach (eDataTypes dataType in dataTypes)
            {
                if (Columns.Count <= i)
                {
                    return;
                }
                Columns[i].DataType = dataType;
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
            if(Columns.Count <= 0)
            {
                CreateColumnFormatList(paddingTypes.Length);
            }
            if (paddingTypes.Length > Columns.Count)
            {
                throw new ArgumentException("paddingTypes has more elements than Columns");
            }
            foreach (PaddingAlignmentType paddingType in paddingTypes)
            {
                if (Columns.Count <= i)
                {
                    return;
                }
                Columns[i].PaddingType = paddingType;
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
            if (Columns.Count <= 0)
            {
                CreateColumnFormatList(UseColumns.Length);
            }
            if (UseColumns.Length > Columns.Count)
            {
                throw new ArgumentException("UseColumns has more elements than Columns");
            }
            foreach (bool UseColumn in UseColumns)
            {
                if (Columns.Count <= i)
                {
                    return;
                }
                Columns[i].UseColumn = UseColumn;
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
            if (Columns.Count <= 0)
            {
                CreateColumnFormatList(Names.Length);
            }
            if (Names.Length > Columns.Count)
            {
                throw new ArgumentException("Names has more elements than Columns");
            }
            foreach (string name in Names)
            {
                if (Columns.Count <= i)
                {
                    return;
                }
                Columns[i].Name = name;
                i++;
            }
            FirstRowIsHeader = true;
        }
    }
}
