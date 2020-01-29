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
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        #region LoadFromDataReader
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None)
        {
            var r = LoadFromDataReader(Reader, PrintHeaders);
            
            int rows = r.Rows - 1;
            if (rows >= 0 && r.Columns > 0)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + (rows <= 0 ? 1 : rows), _fromCol + r.Columns - 1), TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders)
        {
            if (Reader == null)
            {
                throw (new ArgumentNullException("Reader", "Reader can't be null"));
            }
            int fieldCount = Reader.FieldCount;

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    _worksheet.SetValueInner(row, col++, Reader.GetName(i));
                }
                row++;
                col = _fromCol;
            }            
            while (Reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    _worksheet.SetValueInner(row, col++, Reader.GetValue(i));
                }
                row++;
                col = _fromCol;
            }
            return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + fieldCount - 1];
        }
#if !NET35 && !NET40
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None,  CancellationToken? cancellationToken=null)
        {
            cancellationToken = cancellationToken ?? CancellationToken.None;
            var r = await LoadFromDataReaderAsync(Reader, PrintHeaders, cancellationToken.Value).ConfigureAwait(false);

            if (cancellationToken.Value.IsCancellationRequested) return r;

            int rows = r.Rows - 1;
            if (rows >= 0 && r.Columns > 0)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + (rows <= 0 ? 1 : rows), _fromCol + r.Columns - 1), TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders)
        {
            return await LoadFromDataReaderAsync(Reader, PrintHeaders, CancellationToken.None);
        }
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, CancellationToken cancellationToken)
        {
            if (Reader == null)
            {
                throw (new ArgumentNullException("Reader", "Reader can't be null"));
            }
            int fieldCount = Reader.FieldCount;

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    _worksheet.SetValueInner(row, col++, Reader.GetName(i));
                }
                row++;
                col = _fromCol;
            }

            while (await Reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    _worksheet.SetValueInner(row, col++, Reader.GetValue(i));
                }
                row++;
                col = _fromCol;
                if (row % 100 == 0 && cancellationToken.IsCancellationRequested)    //Check every 100 rows
                {
                    return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + fieldCount - 1];
                }
            }
            return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + fieldCount - 1];
        }
#endif
        #endregion
        #region LoadFromDataTable
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles TableStyle)
        {
            var r = LoadFromDataTable(Table, PrintHeaders);

            int rows = (Table.Rows.Count == 0 ? 1 : Table.Rows.Count) + (PrintHeaders ? 1 : 0);
            if (rows >= 0 && Table.Columns.Count > 0)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + rows - 1, _fromCol + Table.Columns.Count - 1), Table.TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders)
        {
            if (Table == null)
            {
                throw (new ArgumentNullException("Table can't be null"));
            }

            if (Table.Rows.Count == 0 && PrintHeaders == false)
            {
                return null;
            }

            //var rowArray = new List<object[]>();
            var row = _fromRow;
            if (PrintHeaders)
            {
                _worksheet._values.SetValueRow_Value(_fromRow, _fromCol, Table.Columns.Cast<DataColumn>().Select((dc) => { return dc.Caption; }).ToArray());
                row++;
            }
            foreach (DataRow dr in Table.Rows)
            {
                _worksheet._values.SetValueRow_Value(row++, _fromCol, dr.ItemArray);
            }
            if (row != _fromRow) row--;
            return _worksheet.Cells[_fromRow, _fromCol, row, _fromCol + Table.Columns.Count - 1];
        }
#endregion
#region LoadFromArrays
        /// <summary>
        /// Loads data from the collection of arrays of objects into the range, starting from
        /// the top-left cell.
        /// </summary>
        /// <param name="Data">The data.</param>
        public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> Data)
        {
            //thanx to Abdullin for the code contribution
            if (Data == null) throw new ArgumentNullException("data");

            var rowArray = new List<object[]>();
            var maxColumn = 0;
            var row = _fromRow;
            foreach (object[] item in Data)
            {
                //rowArray.Add(item);
                _worksheet._values.SetValueRow_Value(row, _fromCol, item);
                if (maxColumn < item.Length) maxColumn = item.Length;
                row++;
            }
            if (rowArray.Count == 0) return null; //Issue #57
            //_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1,
            //    (List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
            //    {
            //        rowIx -= _fromRow;
            //        columnIx -= _fromCol;

            //        var values = ((List<object[]>)value);
            //        if (values.Count <= rowIx) return;
            //        var item = values[rowIx];
            //        if (item.Length <= columnIx) return;

            //        var val = item[columnIx];
            //        if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
            //        {
            //            list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
            //        }
            //    }, rowArray);

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1];
        }
#endregion
#region LoadFromCollection
        /// <summary>
        /// Load a collection into a the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection)
        {
            return LoadFromCollection<T>(Collection, false, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="memberFlags">Property flags to use</param>
        /// <param name="Members">The properties to output. Must be of type T</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle, BindingFlags memberFlags, MemberInfo[] Members)
        {
            var type = typeof(T);
            bool isSameType = true;
            if (Members == null)
            {
                Members = type.GetProperties(memberFlags);
            }
            else
            {
                if (Members.Length == 0)   //Fixes issue 15555
                {
                    throw (new ArgumentException("Parameter Members must have at least one property. Length is zero"));
                }
                foreach (var t in Members)
                {
                    if (t.DeclaringType != null && t.DeclaringType != type)
                    {
                        isSameType = false;
                    }
                    //Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
                    if (t.DeclaringType != null && t.DeclaringType != type && !TypeCompat.IsSubclassOf(type, t.DeclaringType) && !TypeCompat.IsSubclassOf(t.DeclaringType, type))
                    {
                        throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
                    }
                }
            }

            // create buffer
            object[,] values = new object[(PrintHeaders ? Collection.Count() + 1 : Collection.Count()), Members.Count()];

            int col = 0, row = 0;
            if (Members.Length > 0 && PrintHeaders)
            {
                foreach (var t in Members)
                {
                    var descriptionAttribute = t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() as DescriptionAttribute;
                    var header = string.Empty;
                    if (descriptionAttribute != null)
                    {
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        var displayNameAttribute =
                            t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as
                            DisplayNameAttribute;
                        if (displayNameAttribute != null)
                        {
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            header = t.Name.Replace('_', ' ');
                        }
                    }
                    //_worksheet.SetValueInner(row, col++, header);
                    values[row, col++] = header;
                }
                row++;
            }

            if (!Collection.Any() && (Members.Length == 0 || PrintHeaders == false))
            {
                return null;
            }

            foreach (var item in Collection)
            {
                col = 0;
                if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
                {
                    values[row, col++] = item;
                }
                else
                {
                    foreach (var t in Members)
                    {
                        if (isSameType == false && item.GetType().GetMember(t.Name, memberFlags).Length == 0)
                        {
                            col++;
                            continue; //Check if the property exists if and inherited class is used
                        }
                        else if (t is PropertyInfo)
                        {
                            values[row, col++] = ((PropertyInfo)t).GetValue(item, null);
                        }
                        else if (t is FieldInfo)
                        {
                            values[row, col++] = ((FieldInfo)t).GetValue(item);
                        }
                        else if (t is MethodInfo)
                        {
                            values[row, col++] = ((MethodInfo)t).Invoke(item, null);
                        }
                    }
                }
                row++;
            }

            _worksheet.SetRangeValueInner(_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1, values);

            //Must have at least 1 row, if header is showen
            if (row == 1 && PrintHeaders)
            {
                row++;
            }

            var r = _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1];

            if (TableStyle != TableStyles.None)
            {
                var tbl = _worksheet.Tables.Add(r, "");
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }
#endregion
#region LoadFromText
        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// Default settings is Comma separation
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text)
        {
            return LoadFromText(Text, new ExcelTextFormat());
        }
        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format)
        {
            if (string.IsNullOrEmpty(Text))
            {
                var r = _worksheet.Cells[_fromRow, _fromCol];
                r.Value = "";
                return r;
            }

            if (Format == null) Format = new ExcelTextFormat();


            string[] lines;
            if (Format.TextQualifier == 0)
            {
                lines = Regex.Split(Text, Format.EOL);
            }
            else
            {
                lines = GetLines(Text, Format);
            }

            int row = 0;
            int col = 0;
            int maxCol = col;
            int lineNo = 1;
            //var values = new List<object>[lines.Length];
            foreach (string line in lines)
            {
                var items = new List<object>();
                //values[row] = items;

                if (lineNo > Format.SkipLinesBeginning && lineNo <= lines.Length - Format.SkipLinesEnd)
                {
                    col = 0;
                    string v = "";
                    bool isText = false, isQualifier = false;
                    int QCount = 0;
                    int lineQCount = 0;
                    foreach (char c in line)
                    {
                        if (Format.TextQualifier != 0 && c == Format.TextQualifier)
                        {
                            if (!isText && v != "")
                            {
                                throw (new Exception(string.Format("Invalid Text Qualifier in line : {0}", line)));
                            }
                            isQualifier = !isQualifier;
                            QCount += 1;
                            lineQCount++;
                            isText = true;
                        }
                        else
                        {
                            if (QCount > 1 && !string.IsNullOrEmpty(v))
                            {
                                v += new string(Format.TextQualifier, QCount / 2);
                            }
                            else if (QCount > 2 && string.IsNullOrEmpty(v))
                            {
                                v += new string(Format.TextQualifier, (QCount - 1) / 2);
                            }

                            if (isQualifier)
                            {
                                v += c;
                            }
                            else
                            {
                                if (c == Format.Delimiter)
                                {
                                    items.Add(ConvertData(Format, v, col, isText));
                                    v = "";
                                    isText = false;
                                    col++;
                                }
                                else
                                {
                                    if (QCount % 2 == 1)
                                    {
                                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
                                    }
                                    v += c;
                                }
                            }
                            QCount = 0;
                        }
                    }
                    if (QCount > 1 && (v != "" && QCount == 2))
                    {
                        v += new string(Format.TextQualifier, QCount / 2);
                    }
                    if (lineQCount % 2 == 1)
                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
                    items.Add(ConvertData(Format, v, col, isText));

                    _worksheet._values.SetValueRow_Value(_fromRow + row, _fromCol, items);

                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + maxCol];
        }

        private string[] GetLines(string text, ExcelTextFormat Format)
        {
            if (Format.EOL == null || Format.EOL.Length == 0) return new string[] { text };
            var eol = Format.EOL;
            var list = new List<string>();
            var inTQ = false;
            var prevLineStart = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == Format.TextQualifier)
                {
                    inTQ = !inTQ;
                }
                else if (!inTQ)
                {
                    if (IsEOL(text, i, eol))
                    {
                        list.Add(text.Substring(prevLineStart, i - prevLineStart));
                        i += eol.Length - 1;
                        prevLineStart = i + 1;
                    }
                }
            }

            if (inTQ)
            {
                throw (new ArgumentException(string.Format("Text delimiter is not closed in line : {0}", list.Count)));
            }

            if (prevLineStart >= Format.EOL.Length && IsEOL(text, prevLineStart - Format.EOL.Length, Format.EOL))
            {
                //list.Add(text.Substring(prevLineStart- Format.EOL.Length, Format.EOL.Length));
                list.Add("");
            }
            else
            {
                list.Add(text.Substring(prevLineStart));
            }
            return list.ToArray();
        }
        private bool IsEOL(string text, int ix, string eol)
        {
            for (int i = 0; i < eol.Length; i++)
            {
                if (text[ix + i] != eol[i])
                    return false;
            }
            return ix + eol.Length <= text.Length;
        }

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            var r = LoadFromText(Text, Format);

            var tbl = _worksheet.Tables.Add(r, "");
            tbl.ShowHeader = FirstRowIsHeader;
            tbl.TableStyle = TableStyle;

            return r;
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Encoding.ASCII));
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
        }
#region LoadFromText async
#if !NET35 && !NET40
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile)
        {
            var fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            var sr = new StreamReader(fs, Encoding.ASCII);            
            return LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false));
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile, ExcelTextFormat Format)
        {
            var fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            var sr = new StreamReader(fs, Format.Encoding);
            return LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false), Format);
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public async Task<ExcelRangeBase> LoadFromTextAsync(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            var fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            var sr = new StreamReader(fs, Format.Encoding);
            return LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false), Format, TableStyle, FirstRowIsHeader);
        }
#endif
#endregion
#endregion
    }
}
