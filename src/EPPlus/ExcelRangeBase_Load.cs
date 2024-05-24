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
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
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
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <param name="Transpose">Transpose the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, bool Transpose, TableStyles TableStyle = TableStyles.None)
        {
            var r = Transpose ? LoadFromDataReader(Reader, PrintHeaders, Transpose) : LoadFromDataReader(Reader, PrintHeaders);

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
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="Transpose">Must be true to transpose data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, bool Transpose)
        {
            if (Reader == null)
            {
                throw (new ArgumentNullException("Reader", "Reader can't be null"));
            }
            if(Transpose == false)
            {
                throw (new ArgumentNullException("Transpose", "Must be true, use LeadFromDataReader without argument Transpose instead"));
            }
            int fieldCount = Reader.FieldCount;

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    _worksheet.SetValueInner(row++, col, Reader.GetName(i));
                }
                col++;
                row = _fromRow;
            }
            while (Reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    _worksheet.SetValueInner(row++, col, Reader.GetValue(i));
                }
                col++;
                row = _fromRow;
            }
            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + fieldCount - 1, col - 1];
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
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="Transpose"></param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, string TableName, bool Transpose, TableStyles TableStyle = TableStyles.None, CancellationToken? cancellationToken = null)
        {
            cancellationToken = cancellationToken ?? CancellationToken.None;
            var r = await LoadFromDataReaderAsync(Reader, PrintHeaders, cancellationToken.Value, Transpose).ConfigureAwait(false);

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
        /// <param name="Transpose">If the data should be transposed on read or not</param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, bool Transpose)
        {
            return Transpose ? await LoadFromDataReaderAsync(Reader, PrintHeaders, CancellationToken.None, Transpose) : await LoadFromDataReaderAsync(Reader, PrintHeaders, CancellationToken.None);
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
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="cancellationToken">The cancellation token to use</param>
        /// <param name="Transpose"></param>
        /// <returns>The filled range</returns>
        public async Task<ExcelRangeBase> LoadFromDataReaderAsync(DbDataReader Reader, bool PrintHeaders, CancellationToken cancellationToken, bool Transpose)
        {
            if (Reader == null)
            {
                throw (new ArgumentNullException("Reader", "Reader can't be null"));
            }
            if (Transpose == false)
            {
                throw (new ArgumentNullException("Transpose", "Must be true, use LeadFromDataReaderAsync without argument Transpose instead"));
            }
            int fieldCount = Reader.FieldCount;

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    _worksheet.SetValueInner(row++, col, Reader.GetName(i));
                }
                col++;
                row = _fromRow;
            }

            while (await Reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    _worksheet.SetValueInner(row++, col, Reader.GetValue(i));
                }
                col++;
                row = _fromRow;
                if (row % 100 == 0 && cancellationToken.IsCancellationRequested)    //Check every 100 columns
                {
                    return _worksheet.Cells[_fromRow, _fromCol, _fromRow + fieldCount - 1, col - 1 ];
                }
            }
            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + fieldCount - 1, col - 1];
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
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles? TableStyle)
        {
            var parameters = new LoadFromDataTableParams
            {
                PrintHeaders = PrintHeaders,
                TableStyle = TableStyle
            };
            var func = new LoadFromDataTable(this, Table, parameters);
            return func.Load();
        }
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <param name="Transpose">Transpose the loaded data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles? TableStyle, bool Transpose)
        {
            var parameters = new LoadFromDataTableParams
            {
                PrintHeaders = PrintHeaders,
                TableStyle = TableStyle,
                Transpose = Transpose,
            };
            var func = new LoadFromDataTable(this, Table, parameters);
            return func.Load();
        }
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders)
        {
            return LoadFromDataTable(Table, PrintHeaders, null);
        }

        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="table">The datatable to load</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable table)
        {
            return LoadFromDataTable(table, false, null);
        }

        /// <summary>
        /// Load the data from the <see cref="DataTable"/> starting from the top left cell of the range
        /// </summary>
        /// <param name="table"></param>
        /// <param name="paramsConfig"><see cref="Action{LoacFromCollectionParams}"/> to provide parameters to the function</param>
        /// <example>
        /// <code>
        /// sheet.Cells["C1"].LoadFromDataTable(dataTable, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </code>
        /// </example>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable table, Action<LoadFromDataTableParams> paramsConfig)
        {
            var parameters = new LoadFromDataTableParams();
            paramsConfig.Invoke(parameters);
            return LoadFromDataTable(table, parameters.PrintHeaders, parameters.TableStyle, parameters.Transpose);
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
            if (!(Data?.Any() ?? false)) return null;

            var maxColumn = 0;
            var row = _fromRow;
            foreach (object[] item in Data)
            {
                _worksheet._values.SetValueRow_Value(row, _fromCol, item);
                if (maxColumn < item.Length) maxColumn = item.Length;
                row++;
            }

            return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + maxColumn - 1];
        }
        /// <summary>
        /// Loads data from the collection of arrays of objects into the range transposed, starting from
        /// the top-left cell.
        /// </summary>
        /// <param name="Data"></param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromArraysTransposed(IEnumerable<object[]> Data)
        {
            //thanx to Abdullin for the code contribution
            if (!(Data?.Any() ?? false)) return null;

            var maxRow = 0;
            var col = _fromCol;
            foreach (object[] item in Data)
            {
                _worksheet._values.SetValueRow_ValueTransposed(_fromRow, col, item);
                if (maxRow < item.Length) maxRow = item.Length;
                col++;
            }

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + maxRow - 1, col - 1];
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
            var type = typeof(T);
            var attr = type.GetFirstAttributeOfType<EpplusTableAttribute>();
            if(attr != null)
            {
                var range = LoadFromCollection(Collection, attr.PrintHeaders, attr.TableStyle, BindingFlags.Public | BindingFlags.Instance, null);
                if(attr.AutofitColumns)
                {
                    range.AutoFitColumns();
                }
                if(attr.AutoCalculate)
                {
                    range.Calculate();
                }
                return range;
            }
            return LoadFromCollection<T>(Collection, false, null, BindingFlags.Public | BindingFlags.Instance, null);
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
            return LoadFromCollection<T>(Collection, PrintHeaders, null, BindingFlags.Public | BindingFlags.Instance, null);
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
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="Transpose">Will load data transposed</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle, bool Transpose)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, Transpose, BindingFlags.Public | BindingFlags.Instance, null);
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
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle, BindingFlags memberFlags, MemberInfo[] Members)
        {
            return LoadFromCollectionInternal(Collection, PrintHeaders, TableStyle, memberFlags, Members);
        }

        private ExcelRangeBase LoadFromCollectionInternal<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle, BindingFlags memberFlags, MemberInfo[] Members)
        {
            if (Collection is IEnumerable<IDictionary<string, object>>)
            {
                if (Members == null)
                    return LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, PrintHeaders, TableStyle);
                return LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, PrintHeaders, TableStyle, Members.Select(x => x.Name));
            }
            var param = new LoadFromCollectionParams
            {
                PrintHeaders = PrintHeaders,
                TableStyle = TableStyle,
                BindingFlags = memberFlags,
                Members = Members
            };
            var func = new LoadFromCollection<T>(this, Collection, param);
            return func.Load();
        }
        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="Transpose">Will insert data transposed</param>
        /// <param name="memberFlags">Property flags to use</param>
        /// <param name="Members">The properties to output. Must be of type T</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle, bool Transpose, BindingFlags memberFlags, MemberInfo[] Members)
        {
            return LoadFromCollectionInternal(Collection, PrintHeaders, TableStyle, Transpose, memberFlags, Members);
        }

        private ExcelRangeBase LoadFromCollectionInternal<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles? TableStyle, bool Transpose, BindingFlags memberFlags, MemberInfo[] Members)
        {
            if (Collection is IEnumerable<IDictionary<string, object>>)
            {
                if (Members == null)
                    return LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, c =>
                    {
                        c.PrintHeaders = PrintHeaders;
                        c.TableStyle = TableStyle;
                        c.Transpose = Transpose;
                    });
                return LoadFromDictionaries(Collection as IEnumerable<IDictionary<string, object>>, c =>
                {
                    c.PrintHeaders = PrintHeaders;
                    c.TableStyle = TableStyle;
                    c.Transpose = Transpose;
                    c.SetKeys(Members.Select(x => x.Name).ToArray());
                });
            }
            var param = new LoadFromCollectionParams
            {
                PrintHeaders = PrintHeaders,
                TableStyle = TableStyle,
                BindingFlags = memberFlags,
                Members = Members,
                Transpose = Transpose,
            };
            var func = new LoadFromCollection<T>(this, Collection, param);
            return func.Load();
        }

        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="collection">The collection to load</param>
        /// <param name="paramsConfig"><see cref="Action{LoacFromCollectionParams}"/> to provide parameters to the function</param>
        /// <example>
        /// <code>
        /// sheet.Cells["C1"].LoadFromCollection(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </code>
        /// </example>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> collection, Action<LoadFromCollectionParams> paramsConfig)
        {
            var param = new LoadFromCollectionParams();
            paramsConfig.Invoke(param);
            if (collection is IEnumerable<IDictionary<string, object>>)
            {
                if (param.Members == null)
                {
                    return LoadFromDictionaries(collection as IEnumerable<IDictionary<string, object>>, c =>
                    {
                        c.PrintHeaders = param.PrintHeaders;
                        c.TableStyle = param.TableStyle;
                        c.Transpose = param.Transpose;
                    });
                }
                return LoadFromDictionaries(collection as IEnumerable<IDictionary<string, object>>, c =>
                {
                    c.PrintHeaders = param.PrintHeaders;
                    c.TableStyle = param.TableStyle;
                    c.Transpose = param.Transpose;
                    c.SetKeys(param.Members.Select(x => x.Name).ToArray());
                });
                //return LoadFromDictionaries(collection as IEnumerable<IDictionary<string, object>>, param.PrintHeaders, param.TableStyle, param.Members.Select(x => x.Name));
            }
            var func = new LoadFromCollection<T>(this, collection, param);
            return func.Load();
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

            return LoadFromTextPrivate(Text, Format, Format.TableStyle, Format.FirstRowIsHeader);
        }

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style. If this parameter is not null no table will be created.</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles? TableStyle, bool FirstRowIsHeader)
        {
            return LoadFromTextPrivate(Text, Format, TableStyle, FirstRowIsHeader);
        }

        private ExcelRangeBase LoadFromTextPrivate(string Text, ExcelTextFormat Format, TableStyles? TableStyle, bool FirstRowIsHeader)
        {
            var parameters = new LoadFromTextParams
            {
                Format = Format
            };
            var func = new LoadFromText(this, Text, parameters);
            var r = func.Load();

            if (r != null && TableStyle.HasValue)
            {
                var tbl = _worksheet.Tables.Add(r, "");
                tbl.ShowHeader = FirstRowIsHeader;
                tbl.TableStyle = TableStyle.Value;
            }
            return r;
        }

        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell using ASCII Encoding.
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
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

            return LoadFromTextPrivate(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, Format.TableStyle,  Format.FirstRowIsHeader);
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles? TableStyle, bool FirstRowIsHeader)
        {
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
        }

        /// <summary>
        /// Loads a fixed width text file into range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text file</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormatFixedWidth Format)
        {
            if (string.IsNullOrEmpty(Text))
            {
                var r = _worksheet.Cells[_fromRow, _fromCol];
                r.Value = "";
                return r;
            }
            var func = new LoadFromFixedWidthText(this, Text, Format);
            var range =  func.Load();
            if (range != null && Format.TableStyle.HasValue)
            {
                var tbl = _worksheet.Tables.Add(range, "");
                tbl.ShowHeader = Format.FirstRowIsHeader;
                tbl.TableStyle = Format.TableStyle.Value;
            }
            return range;
        }

        /// <summary>
        /// Loads a fixed width text file into range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormatFixedWidth Format)
        {
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
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
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

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
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

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
            if (TextFile.Exists == false)
            {
                throw (new ArgumentException($"File does not exist {TextFile.FullName}"));
            }

            var fs = new FileStream(TextFile.FullName, FileMode.Open, FileAccess.Read);
            var sr = new StreamReader(fs, Format.Encoding);
            return LoadFromText(await sr.ReadToEndAsync().ConfigureAwait(false), Format, TableStyle, FirstRowIsHeader);
        }
#endif
        #endregion
        #endregion
        #region LoadFromDictionaries
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items)
        {
            return LoadFromDictionaries(items, false, TableStyles.None, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders)
        {
            return LoadFromDictionaries(items, printHeaders, TableStyles.None, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders, TableStyles? tableStyle)
        {
            return LoadFromDictionaries(items, printHeaders, tableStyle, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries</param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="keys">Keys that should be used, keys omitted will not be included</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders, TableStyles? tableStyle, IEnumerable<string> keys)
        {
            var param = new LoadFromDictionariesParams
            {
                PrintHeaders = printHeaders,
                TableStyle = tableStyle
            };
            if(keys != null && keys.Any())
            {
                param.SetKeys(keys.ToArray());
            }
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }
#if !NET35 && !NET40
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries</param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="keys">Keys that should be used, keys omitted will not be included</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<dynamic> items, bool printHeaders, TableStyles? tableStyle, IEnumerable<string> keys)
        {
            var param = new LoadFromDictionariesParams
            {
                PrintHeaders = printHeaders,
                TableStyle = tableStyle
            };
            if (keys != null && keys.Any())
            {
                param.SetKeys(keys.ToArray());
            }
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }
#endif

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/ExpandoObjects</param>
        /// <param name="paramsConfig"><see cref="Action{LoadFromDictionariesParams}"/> to provide parameters to the function</param>
        /// <example>
        /// sheet.Cells["C1"].LoadFromDictionaries(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, Action<LoadFromDictionariesParams> paramsConfig)
        {
            var param = new LoadFromDictionariesParams();
            paramsConfig.Invoke(param);
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }

#if !NET35 && !NET40
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/ExpandoObjects</param>
        /// <param name="paramsConfig"><see cref="Action{LoadFromDictionariesParams}"/> to provide parameters to the function</param>
        /// <example>
        /// sheet.Cells["C1"].LoadFromDictionaries(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<dynamic> items, Action<LoadFromDictionariesParams> paramsConfig)
        {
            var param = new LoadFromDictionariesParams();
            paramsConfig.Invoke(param);
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }        
#endif

        #endregion
    }
}
