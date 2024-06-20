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
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Data;
using System.Collections.Generic;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Export.ToCollection;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;


#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        #region ToDataTable

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the <see cref="ToDataTableOptions.Default"/> settings.
        /// </summary>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable()
        {
            return ToDataTable(ToDataTableOptions.Default);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="configHandler">Configures the settings used to convert the range.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler)
        {
            var o = ToDataTableOptions.Default;
            configHandler.Invoke(o);
            return ToDataTable(o);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="options">Sets the settings used to convert the range.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(ToDataTableOptions options)
        {
            var func = new ToDataTable(options, this);
            return func.Execute();
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="configHandler">Configures the settings used to convert the range.</param>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler, DataTable dataTable)
        {
            var o = ToDataTableOptions.Default;
            configHandler.Invoke(o);
            return ToDataTable(o, dataTable);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(DataTable dataTable)
        {
            return ToDataTable(ToDataTableOptions.Default, dataTable);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="options">Sets the settings used to convert the range.</param>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(ToDataTableOptions options, DataTable dataTable)
        {
            var func = new ToDataTable(options, this);
            return func.Execute(dataTable);
        }
        #endregion
        #region ToText / SaveToText
        /// <summary>
        /// Converts a range to text in CSV format.
        /// </summary>
        /// <returns>A string containing the text</returns>
        public string ToText()
        {
            return ToText(new ExcelOutputTextFormat());
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the csv text</param>
        /// <returns>A string containing the text</returns>
        public string ToText(ExcelOutputTextFormat Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                SaveToText(ms, Format);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return sr.ReadToEnd();
            }
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public void SaveToText(FileInfo file, ExcelOutputTextFormat Format)
        {
            using (var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write))
            {
				SaveToText(fileStream, Format);
                fileStream.Close();
			}
		}
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public void SaveToText(Stream stream, ExcelOutputTextFormat Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormat();
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) sw.Write(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;
            var fromCol = _fromCol;
            var toCol = _toCol;
            var fromRow = _fromRow;
            var toRow = _toRow;
            if (Format.DataIsTransposed)
            {
                fromRow = _fromCol;
                toRow = _toCol;
                fromCol = _fromRow;
                toCol = _toRow;
            }
            bool hasTextQ = Format.TextQualifier != '\0';
            string doubleTextQualifiers = new string(Format.TextQualifier, 2);
            var skipLinesBegining = Format.SkipLinesBeginning + (Format.FirstRowIsHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = fromRow; row <= toRow; row++)
            {
                if (row == fromRow && Format.FirstRowIsHeader)
                {
                    sw.Write(WriteHeaderRow(Format, hasTextQ, row, ci));
                    continue;
                }

                if (SkipLines(Format, row, skipLinesBegining))
                {
                    continue;
                }
                string finalRow = "";
                for (int col = fromCol; col <= toCol; col++)
                {
                    string t = Format.DataIsTransposed ? GetTextCSV(Format, maxFormats, ci, col, row, out bool isText) : GetTextCSV(Format, maxFormats, ci, row, col, out isText);

                    if (hasTextQ && isText)
                    {
                        //sw.Write(Format.TextQualifier);
                        //sw.Write(t.Replace(Format.TextQualifier.ToString(), doubleTextQualifiers));
                        //sw.Write(Format.TextQualifier);
                        finalRow += Format.TextQualifier;
                        finalRow += t.Replace(Format.TextQualifier.ToString(), doubleTextQualifiers);
                        finalRow += Format.TextQualifier;
                    }
                    else
                    {
                        finalRow += t;
                    }
                    if (col != toCol) finalRow += Format.Delimiter;//sw.Write(Format.Delimiter);
                }
                if (Format.ShouldUseRow != null && Format.ShouldUseRow.Invoke(finalRow) == false)
                {
                    continue;
                }
                else
                {
                    sw.Write(finalRow);
                }
                if (row != toRow - Format.SkipLinesEnd) sw.Write(Format.EOL);
            }
            if (!string.IsNullOrEmpty(Format.Footer)) sw.Write(Format.EOL + Format.Footer);
            sw.Flush();
        }

        /// <summary>
        /// Converts a range to text in Fixed Width format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the Fixed Width text</param>
        /// <returns>A string containing the text</returns>
        public string ToText(ExcelOutputTextFormatFixedWidth Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                SaveToText(ms, Format);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return sr.ReadToEnd();
            }
        }

        /// <summary>
        /// Converts a range to text in fixed widths format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the fixed width text</param>
        public void SaveToText(FileInfo file, ExcelOutputTextFormatFixedWidth Format)
        {
            using (var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write))
            {
                SaveToText(fileStream, Format);
                fileStream.Close();
            }
        }

        /// <summary>
        /// Converts a range to text in Fixed Width format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the fixed width text</param>
        public void SaveToText(Stream stream, ExcelOutputTextFormatFixedWidth Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormatFixedWidth();
            if(Format.Columns == null) throw new ArgumentNullException("Format.ColumnFormat: Set ColumnFormat.Length or ColumnFormat.Position");
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) sw.Write(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;

            var skipLinesBegining = Format.SkipLinesBeginning + (Format.ExcludeHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = _fromRow; row <= _toRow; row++)
            {
                if (row == _fromRow && Format.ExcludeHeader)
                {
                    continue;
                }
                if (skipLinesBegining > row - _fromRow || Format.SkipLinesEnd > _toRow - row)
                {
                    continue;
                }
                string fc = "";
                int ix = 0;
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    if (Format.Columns[ix].UseColumn)
                    {
                        string text;
                        eDataTypes dataType;
                        if ( string.IsNullOrEmpty( Format.Columns[ix].Name))
                        {
                            text = GetTextFixedWidth(Format, maxFormats, ci, row, col, out dataType);
                        }
                        else
                        {
                            text = Format.Columns[ix].Name;
                            dataType = eDataTypes.String;
                        }
                        var padding = 0;
                        if (Format.ReadType == FixedWidthReadType.Length)
                        {
                            padding = Format.Columns[ix].Length - text.Length;
                        }
                        else if (Format.ReadType == FixedWidthReadType.Positions)
                        {
                            if (ix+1 < Format.Columns.Count)
                            {
                                padding = (Format.Columns[ix + 1].Position - Format.Columns[ix].Position) - text.Length;
                            }
                            else if (Format.Columns[ix].Length > 0)
                            {
                                padding = Format.Columns[ix].Length - text.Length;
                            }
                        }
                        if (padding > 0)
                        {
                            PaddingAlignmentType pat = Format.Columns[ix].PaddingType;
                            bool numericPadding = (dataType==eDataTypes.Number || dataType == eDataTypes.Percent) && double.TryParse(text, NumberStyles.Any, Format.Culture, out double result);
                            if (Format.Columns[ix].PaddingType == PaddingAlignmentType.Auto)
                            {
                                if ( numericPadding || text.EndsWith("%"))
                                {
                                    pat = PaddingAlignmentType.Right;
                                }
                                else
                                {
                                    pat = PaddingAlignmentType.Left;
                                }
                            }
                            for (int i = 0; i < padding; i++)
                            {
                                if (pat == PaddingAlignmentType.Left)
                                {
                                    if (numericPadding && Format.PaddingCharacterNumeric != null)
                                    {
                                        text += Format.PaddingCharacterNumeric;
                                    }
                                    else
                                    {
                                        text += Format.PaddingCharacter;
                                    }
                                }
                                else if (pat == PaddingAlignmentType.Right)
                                {
                                    if (numericPadding && Format.PaddingCharacterNumeric != null)
                                    {
                                        text = Format.PaddingCharacterNumeric + text;
                                    }
                                    else
                                    {
                                        text = Format.PaddingCharacter + text;
                                    }
                                }
                            }
                        }
                        else if (padding < 0)
                        {
                            if (Format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.Truncate)
                            {
                                if (Format.ReadType == FixedWidthReadType.Length)
                                {
                                    text = text.Substring(0, Format.Columns[ix].Length);
                                }
                                else if (Format.ReadType == FixedWidthReadType.Positions)
                                {
                                    if(Format.Columns[ix].Length > 0)
                                    {
                                        text = text.Substring(0, Format.Columns[ix].Length);
                                    }
                                    else
                                    {
                                        text = text.Substring(0);
                                    }
                                }
                            }
                            else
                            {
                                throw new FormatException("String was " + text.Length + ", Expected length of " + Format.Columns[ix].Length);
                            }
                        }
                        if (Format.ReadType == FixedWidthReadType.Positions && Format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.Truncate) //truncate
                        {
                            fc = fc.Substring(0, Format.Columns[ix].Position);
                        }
                        fc += text;
                    }
                    ix++;
                }
                if (Format.ShouldUseRow != null && Format.ShouldUseRow.Invoke(fc) == false)
                {
                    continue;
                }
                else
                {
                    sw.Write(fc + Format.EOL);
                }
            }
            if (!string.IsNullOrEmpty(Format.Footer)) sw.Write(Format.Footer);
            sw.Flush();
        }


        #endregion
        #region ToText / SaveToText async
#if !NET35 && !NET40
        /// <summary>
        /// Converts a range to text in CSV format.
        /// </summary>
        /// <returns>A string containing the text</returns>
        public async Task<string> ToTextAsync()
        {
            ExcelOutputTextFormat Format = new ExcelOutputTextFormat();
            return await ToTextAsync(Format).ConfigureAwait(false);
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the csv text</param>
        /// <returns>A string containing the text</returns>
        public async Task<string> ToTextAsync(ExcelOutputTextFormat Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await SaveToTextAsync(ms, Format).ConfigureAwait(false);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return await sr.ReadToEndAsync().ConfigureAwait(false);
            }
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public async Task SaveToTextAsync(FileInfo file, ExcelOutputTextFormat Format)
        {
            using (var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write))
            {
                await SaveToTextAsync(fileStream, Format).ConfigureAwait(false);
                fileStream.Close();
            }
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public async Task SaveToTextAsync(Stream stream, ExcelOutputTextFormat Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormat();
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) sw.Write(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;
            var fromCol = _fromCol;
            var toCol = _toCol;
            var fromRow = _fromRow;
            var toRow = _toRow;
            if (Format.DataIsTransposed)
            {
                fromRow = _fromCol;
                toRow = _toCol;
                fromCol = _fromRow;
                toCol = _toRow;
            }
            bool hasTextQ = Format.TextQualifier != '\0';
            string encodedTextQualifier = "";
            if (hasTextQ)
            {
                if (Format.EncodedTextQualifiers == null)
                {
                    encodedTextQualifier = new string(Format.TextQualifier, 2);
                }
                else
                {
                    encodedTextQualifier = Format.EncodedTextQualifiers;
                }
            }
            var skipLinesBegining = Format.SkipLinesBeginning + (Format.FirstRowIsHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = fromRow; row <= toRow; row++)
            {
                if (row == fromRow && Format.FirstRowIsHeader)
                {
                    await sw.WriteAsync(WriteHeaderRow(Format, hasTextQ, row, ci)).ConfigureAwait(false);
                    continue;
                }

                if (SkipLines(Format, row, skipLinesBegining))
                {
                    continue;
                }

                for (int col = fromCol; col <= toCol; col++)
                {
                    string t = Format.DataIsTransposed ? GetText(Format, maxFormats, ci, col, row, out eDataTypes dataType) : GetText(Format, maxFormats, ci, row, col, out dataType);

                    if (hasTextQ && dataType==eDataTypes.String)
                    {
                        await sw.WriteAsync(Format.TextQualifier + t.Replace(Format.TextQualifier.ToString(), encodedTextQualifier) + Format.TextQualifier).ConfigureAwait(false);
                    }
                    else
                    {
                        await sw.WriteAsync(t).ConfigureAwait(false);
                    }
                    if (col != toCol) await sw.WriteAsync(Format.Delimiter).ConfigureAwait(false);
                }
                if (row != toRow - Format.SkipLinesEnd) await sw.WriteAsync(Format.EOL).ConfigureAwait(false);
            }
            if (!string.IsNullOrEmpty(Format.Footer)) await sw.WriteAsync(Format.EOL + Format.Footer).ConfigureAwait(false);
            sw.Flush();
        }

        /// <summary>
        /// Converts a range to text in Fixed Width format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the Fixed Width text</param>
        /// <returns>A string containing the text</returns>
        public async Task<string> ToTextAsync(ExcelOutputTextFormatFixedWidth Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await SaveToTextAsync(ms, Format);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return sr.ReadToEnd();
            }
        }

        /// <summary>
        /// Converts a range to text in fixed widths format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the fixed width text</param>
        public async Task SaveToTextAsync(FileInfo file, ExcelOutputTextFormatFixedWidth Format)
        {
            using (var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write))
            {
                await SaveToTextAsync(fileStream, Format);
                fileStream.Close();
            }
        }

        /// <summary>
        /// Converts a range to text in Fixed Width format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the fixed width text</param>
        public async Task SaveToTextAsync(Stream stream, ExcelOutputTextFormatFixedWidth Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormatFixedWidth();
            if (Format.Columns == null) throw new ArgumentNullException("Format.ColumnFormat: Set ColumnFormat.Length or ColumnFormat.Position");
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) await sw.WriteAsync(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;

            var skipLinesBegining = Format.SkipLinesBeginning + (Format.ExcludeHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = _fromRow; row <= _toRow; row++)
            {
                if (row == _fromRow && Format.ExcludeHeader)
                {
                    continue;
                }
                if (skipLinesBegining > row - _fromRow || Format.SkipLinesEnd > _toRow - row)
                {
                    continue;
                }
                string fc = "";
                int ix = 0;
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    if (Format.Columns[ix].UseColumn)
                    {
                        string text;
                        eDataTypes dataType;
                        if (string.IsNullOrEmpty(Format.Columns[ix].Name))
                        {
                            text = GetTextFixedWidth(Format, maxFormats, ci, row, col, out dataType);
                        }
                        else
                        {
                            text = Format.Columns[ix].Name;
                            dataType = eDataTypes.String;
                        }
                        var padding = 0;
                        if (Format.ReadType == FixedWidthReadType.Length)
                        {
                            padding = Format.Columns[ix].Length - text.Length;
                        }
                        else if (Format.ReadType == FixedWidthReadType.Positions)
                        {
                            if (ix + 1 < Format.Columns.Count)
                            {
                                padding = (Format.Columns[ix + 1].Position - Format.Columns[ix].Position) - text.Length;
                            }
                            else if (Format.Columns[ix].Length > 0)
                            {
                                padding = Format.Columns[ix].Length - text.Length;
                            }
                        }
                        if (padding > 0)
                        {
                            PaddingAlignmentType pat = Format.Columns[ix].PaddingType;
                            bool numericPadding = (dataType == eDataTypes.Number || dataType == eDataTypes.Percent) && double.TryParse(text, NumberStyles.Any, Format.Culture, out double result);
                            if (Format.Columns[ix].PaddingType == PaddingAlignmentType.Auto)
                            {
                                if (numericPadding || text.EndsWith("%"))
                                {
                                    pat = PaddingAlignmentType.Right;
                                }
                                else
                                {
                                    pat = PaddingAlignmentType.Left;
                                }
                            }
                            for (int i = 0; i < padding; i++)
                            {
                                if (pat == PaddingAlignmentType.Left)
                                {
                                    if (numericPadding && Format.PaddingCharacterNumeric != null)
                                    {
                                        text += Format.PaddingCharacterNumeric;
                                    }
                                    else
                                    {
                                        text += Format.PaddingCharacter;
                                    }
                                }
                                else if (pat == PaddingAlignmentType.Right)
                                {
                                    if (numericPadding && Format.PaddingCharacterNumeric != null)
                                    {
                                        text = Format.PaddingCharacterNumeric + text;
                                    }
                                    else
                                    {
                                        text = Format.PaddingCharacter + text;
                                    }
                                }
                            }
                        }
                        else if (padding < 0)
                        {
                            if (Format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.Truncate)
                            {
                                if (Format.ReadType == FixedWidthReadType.Length)
                                {
                                    text = text.Substring(0, Format.Columns[ix].Length);
                                }
                                else if (Format.ReadType == FixedWidthReadType.Positions)
                                {
                                    if (Format.Columns[ix].Length > 0)
                                    {
                                        text = text.Substring(0, Format.Columns[ix].Length);
                                    }
                                    else
                                    {
                                        text = text.Substring(0);
                                    }
                                }
                            }
                            else
                            {
                                throw new FormatException("String was " + text.Length + ", Expected length of " + Format.Columns[ix].Length);
                            }
                        }
                        if (Format.ReadType == FixedWidthReadType.Positions && Format.FormatErrorStrategy == FixedWidthFormatErrorStrategy.Truncate) //truncate
                        {
                            fc = fc.Substring(0, Format.Columns[ix].Position);
                        }
                        fc += text;
                    }
                    ix++;
                }
                if (Format.ShouldUseRow != null && Format.ShouldUseRow.Invoke(fc) == false)
                {
                    continue;
                }
                else
                {
                    sw.Write(fc + Format.EOL);
                }
            }
            if (!string.IsNullOrEmpty(Format.Footer)) sw.Write(Format.Footer);
            sw.Flush();
        }

#endif
        #endregion
        #region ToJson
        /// <summary>
        /// Returns the range as JSON
        /// </summary>
        /// <returns>A JSON string</returns>
        public string ToJson()
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            using (var ms = RecyclableMemory.GetStream())
            {
                re.Export(ms);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
        /// <summary>
        /// Returns the range as JSON
        /// </summary>
        /// <param name="settings">Configures settings for the JSON export</param>
        /// <returns></returns>
        public string ToJson(Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            using (var ms = RecyclableMemory.GetStream())
            {
                re.Export(ms);
                return s.Encoding.GetString(ms.ToArray());
            }
        }
        /// <summary>
        /// Saves the range as JSON to a stream.
        /// </summary>
        /// <param name="stream">The writable stream to write the JSON to.</param>
        public void SaveToJson(Stream stream)
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            re.Export(stream);
        }
        /// <summary>
        /// Saves the range as JSON to a stream.
        /// </summary>
        /// <param name="stream">The writable stream to write the JSON to</param>
        /// <param name="settings">Configures settings for the JSON export</param>
        public void SaveToJson(Stream stream, Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            re.Export(stream);
        }
        #endregion
        #region SaveToJson Async
#if !NET35 && !NET40
        /// <summary>
        /// Save the range to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream)
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            await re.ExportAsync(stream);
        }
        /// <summary>
        /// Save the range to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <param name="settings">Settings for the json output.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream, Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            await re.ExportAsync(stream);
        }
#endif
        #endregion
        private static CultureInfo GetCultureInfo(ExcelOutputTextFormat Format)
        {
            var ci = (CultureInfo)(Format.Culture.Clone() ?? CultureInfo.InvariantCulture.Clone());
            if (Format.DecimalSeparator != null)
            {
                ci.NumberFormat.NumberDecimalSeparator = Format.DecimalSeparator;
            }
            if (Format.ThousandsSeparator != null)
            {
                ci.NumberFormat.NumberGroupSeparator = Format.ThousandsSeparator;
            }

            return ci;
        }

        private static CultureInfo GetCultureInfo(ExcelOutputTextFormatFixedWidth Format)
        {
            var ci = (CultureInfo)(Format.Culture.Clone() ?? CultureInfo.InvariantCulture.Clone());
            if (Format.DecimalSeparator != null)
            {
                ci.NumberFormat.NumberDecimalSeparator = Format.DecimalSeparator;
            }
            if (Format.ThousandsSeparator != null)
            {
                ci.NumberFormat.NumberGroupSeparator = Format.ThousandsSeparator;
            }

            return ci;
        }

        private bool SkipLines(ExcelOutputTextFormat Format, int row, int skipLinesBegining)
        {
            var fromRow = _fromRow;
            var toRow = _toRow;
            if (Format.DataIsTransposed)
            {
                fromRow = _fromCol;
                toRow = _toCol;
            }
            return skipLinesBegining > row - fromRow ||
                               Format.SkipLinesEnd > toRow - row;
        }

        private string GetTextCSV(ExcelOutputTextFormat Format, int maxFormats, CultureInfo ci, int row, int col, out bool isText)
        {
            string t = GetText(Format, maxFormats, ci, row, col, out eDataTypes dataType);
            //If a formatted numeric/date value contains the delimitter or a text qualifier treat it as text.
            if (dataType != eDataTypes.String && string.IsNullOrEmpty(t) == false && t.IndexOfAny(new[] { Format.Delimiter, Format.TextQualifier }) >= 0)
            {
                isText = true; 
            }
            else
            {
                isText = dataType == eDataTypes.String;
            }
            return t;
        }

        private string GetTextFixedWidth(ExcelOutputTextFormatFixedWidth format, int maxFormats, CultureInfo ci, int row, int col, out eDataTypes dataType)
        {
            var f = new ExcelOutputTextFormat()
            {
                Header = format.Header,
                Footer = format.Footer,
                FirstRowIsHeader = format.ExcludeHeader,
                UseCellFormat = format.UseCellFormat,
                Formats = format.Formats,
                DecimalSeparator = format.DecimalSeparator,
                ThousandsSeparator = format.ThousandsSeparator,
                EncodedTextQualifiers = format.EncodedTextQualifiers,
            };
            string t = GetText(f, maxFormats, ci, row, col, out dataType);
            if (format.UseTrailingMinus && dataType == eDataTypes.Number)
            {
                double d = Convert.ToDouble(t, format.Culture);
                if (d < 0d)
                {
                    var s = t.Substring(1);
                    t = s + "-";
                }
            }
            return t;
        }

        private string GetText(ExcelOutputTextFormat Format, int maxFormats, CultureInfo ci, int row, int col, out eDataTypes dataType)
        {
            var v = GetCellStoreValue(row, col);

            var ix = col - _fromCol;
            string fmt;
            if (ix < maxFormats)
            {
                fmt = Format.Formats[ix];
            }
            else
            {
                fmt = "";
            }
            string t;

            if (string.IsNullOrEmpty(fmt))
            {
                if (Format.UseCellFormat)
                {
                    t = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);
                    if (ConvertUtil.IsNumericOrDate(v._value))
                    {
                        if(ConvertUtil.IsNumeric(v._value))
                        {
                            dataType = eDataTypes.Number;
                        }
                        else
                        {
                            dataType = eDataTypes.DateTime;
                        }
                    }
                    else
                    {
                        dataType = eDataTypes.String;
                    }
                }
                else
                {
                    if (ConvertUtil.IsNumeric(v._value))
                    {
                        t = ConvertUtil.GetValueDouble(v._value).ToString("r", ci);
                        dataType = eDataTypes.Number;
                    }
                    else if (v._value is DateTime date)
                    {
                        t = date.ToString("G", ci);
                        dataType = eDataTypes.DateTime;
                    }
                    else
                    {
                        t = v._value.ToString();
                        dataType = eDataTypes.String;
                    }
                }
            }
            else
            {
                if (fmt == "$")
                {
                    if (Format.UseCellFormat)
                    {
                        t = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);
                    }
                    else
                    {
                        t = v._value.ToString();
                    }
                    dataType = eDataTypes.String;
                }
                else if (ConvertUtil.IsNumeric(v._value))
                {
                    t = ConvertUtil.GetValueDouble(v._value).ToString(fmt, ci);
                    dataType = eDataTypes.Number;
                }
                else if (v._value is DateTime date)
                {
                    t = date.ToString(fmt, ci);
                    dataType = eDataTypes.DateTime;
                }
                else if (v._value is TimeSpan ts)
                {
#if NET35
                            t = ts.Ticks.ToString(ci);
#else
                    t = ts.ToString(fmt, ci);
#endif
                    dataType = eDataTypes.DateTime;
                }
                else
                {
                    t = v._value.ToString();
                    dataType = eDataTypes.String;
                }
            }
            return t;
        }

        private Core.CellStore.ExcelValue GetCellStoreValue(int row, int col)
        {
            var v = _worksheet.GetCoreValueInner(row, col);
            if (_worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
            {
                var rt = (ExcelRichTextCollection)v._value;
                v._value = rt.Text;
            }
            return v;
        }

        private string WriteHeaderRow(ExcelOutputTextFormat Format, bool hasTextQ, int row, CultureInfo ci)
        {
            var fromCol = _fromCol;
            var toCol = _toCol;
            if (Format.DataIsTransposed)
            {
                fromCol = _fromRow;
                toCol = _toRow;
            }
            var sb = new StringBuilder();
            for (int col = fromCol; col <= toCol; col++)
            {
                var v = Format.DataIsTransposed ? GetCellStoreValue(col, row) : GetCellStoreValue(row, col);
                var s = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);

                if (hasTextQ)
                {
                    sb.Append(Format.TextQualifier);
                    sb.Append(s.Replace(Format.TextQualifier.ToString(), new string(Format.TextQualifier, 2)));
                    sb.Append(Format.TextQualifier);
                }
                else
                {
                    sb.Append(s);
                }

                if (col < toCol)
                    sb.Append(Format.Delimiter);
            }
            if (row != _toRow) sb.Append(Format.EOL);
            return sb.ToString();
        }
        /// <summary>
        /// Returns a collection of T for the range. 
        /// If the range contains multiple addresses the first range is used.
        /// The first row must containt the unique headers used as keys in the row dictionary.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollectionWithMappings<T>(Func<ToCollectionRow, T> setRow)
        {
            return ToCollectionWithMappings(setRow, new ToCollectionRangeOptions());
        }
        /// <summary>
        /// Returns a collection of T for the range. 
        /// If the range contains multiple addresses the first range is used.
        /// The first row must contain the unique headers used as keys in the row dictionary.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <param name="options">Configures the settings for the function</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollectionWithMappings<T>(Func<ToCollectionRow, T> setRow, Action<ToCollectionRangeOptions> options)
        {
            var o = new ToCollectionRangeOptions();
            options.Invoke(o);
            return ToCollectionWithMappings(setRow, o);
        }

        /// <summary>
        /// Returns a collection of T for the range. 
        /// If the range contains multiple addresses the first range is used.
        /// The first row must containt the unique headers used as keys in the row dictionary.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <param name="setRow">The call back function to map each row to the item of type T.</param>
        /// <param name="options">Parameters to the function</param>
        /// <returns>A list of T</returns>
        public List<T> ToCollectionWithMappings<T>(Func<ToCollectionRow, T> setRow, ToCollectionRangeOptions options)
        {
            return ToCollectionRange.ToCollection(this, setRow, options);
        }
        /// <summary>
        /// Returns a collection of T for the range. 
        /// If the range contains multiple addresses the first range is used.
        /// The first row must contain the unique headers used as keys in the row dictionary.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to map to</typeparam>
        /// <returns>A list of T</returns>
        public List<T> ToCollection<T>()
        {
            return ToCollection<T>(new ToCollectionRangeOptions() { HeaderRow = 0 });
        }
        /// <summary>
        /// Automatically maps the range to the properties <see cref="T"/> using the headers.
        /// Using this method requires a headers.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to use</typeparam>
        /// <param name="options">Configures the settings for the function</param>
        /// <returns>A list of <see cref="T"/></returns>
        public List<T> ToCollection<T>(Action<ToCollectionRangeOptions> options)
        {
            var o = new ToCollectionRangeOptions();
            options.Invoke(o);
            return ToCollection<T>(o);
        }
        /// <summary>
        /// Automatically maps the range to the properties <see cref="T"/> using the headers.
        /// Using this method requires a headers.
        /// Headers will be mapped to properties using the name or the attributes without white spaces. 
        /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
        /// </summary>
        /// <typeparam name="T">The type to use</typeparam>
        /// <param name="options">Settings for the method</param>
        /// <returns>A list of <see cref="T"/></returns>
        public List<T> ToCollection<T>(ToCollectionRangeOptions options)
        {
            return ToCollectionRange.ToCollection<T>(this, options);
        }
    }
}