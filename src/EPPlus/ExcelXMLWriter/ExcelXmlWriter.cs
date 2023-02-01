using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.ExcelXMLWriter
{
    internal class ExcelXmlWriter
    {
        ExcelWorksheet _ws;
        ExcelPackage _package;
        private Dictionary<int, int> columnStyles = null;

        public ExcelXmlWriter(ExcelWorksheet worksheet, ExcelPackage package)
        {
            _ws = worksheet;
            _package = package;
        }

        internal void WriteNodes(StreamWriter sw, string xml, ref int startOfNode, ref int endOfNode)
        {
            var prefix = _ws.GetNameSpacePrefix();
            //int startOfNode = 0, endOfNode = 0;

            FindNodePositionAndClearItInit(sw, xml, "cols", ref startOfNode, ref endOfNode);
            UpdateColumnData(sw, prefix);

            FindNodePositionAndClearIt(sw, xml, "sheetData", ref startOfNode, ref endOfNode);
            UpdateRowCellData(sw, prefix);

            FindNodePositionAndClearIt(sw, xml, "mergeCells", ref startOfNode, ref endOfNode);
            _ws._mergedCells.CleanupMergedCells();
            if (_ws._mergedCells.Count > 0)
            {
                UpdateMergedCells(sw, prefix);
            }

            if (_ws.GetNode("d:dataValidations") != null)
            {
                FindNodePositionAndClearIt(sw, xml, "dataValidations", ref startOfNode, ref endOfNode);
                UpdateDataValidation(sw, prefix);
            }
        }

        internal void FindNodePositionAndClearItInit(StreamWriter sw, string xml, string nodeName,
            ref int start, ref int end)
        {
            start = end;
            GetBlockPos(xml, nodeName, ref start, ref end);

            sw.Write(xml.Substring(0, start));
        }

        internal void FindNodePositionAndClearIt(StreamWriter sw, string xml, string nodeName,
            ref int start, ref int end)
        {
            //var xml = _worksheetXml.OuterXml;
            //int startNew = start, endNew = end;

            //The start of a new node is always same as end of last node
            start = end;
            GetBlockPos(xml, nodeName, ref start, ref end);

            //sw.Write(xml.Substring(end, start - end));
            //write a length of zero to erase the temporary node
            sw.Write(xml.Substring(start, 0));

            //int cellStart = colEnd, cellEnd = colEnd;
            //GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);

            //sw.Write(xml.Substring(colEnd, cellStart - colEnd));
            //UpdateRowCellData(sw, prefix);
        }

        private void GetBlockPos(string xml, string tag, ref int start, ref int end)
        {
            Match startmMatch, endMatch;
            startmMatch = Regex.Match(xml.Substring(start), string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

            if (!startmMatch.Success) //Not found
            {
                start = -1;
                end = -1;
                return;
            }
            var startPos = startmMatch.Index + start;
            if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
            {
                end = startPos + startmMatch.Length;
            }
            else
            {
                endMatch = Regex.Match(xml.Substring(start), string.Format("(</[^>]*{0}[^>]*>)", tag));
                if (endMatch.Success)
                {
                    end = endMatch.Index + endMatch.Length + start;
                }
            }
            start = startPos;
        }

        /// <summary>
        /// Inserts the cols collection into the XML document
        /// </summary>
        private void UpdateColumnData(StreamWriter sw, string prefix)
        {
            var cse = new CellStoreEnumerator<ExcelValue>(_ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
            bool first = true;
            while (cse.Next())
            {
                if (first)
                {
                    sw.Write($"<{prefix}cols>");
                    first = false;
                }
                var col = cse.Value._value as ExcelColumn;
                ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;

                sw.Write($"<{prefix}col min=\"{col.ColumnMin}\" max=\"{col.ColumnMax}\"");
                if (col.Hidden == true)
                {
                    sw.Write(" hidden=\"1\"");
                }
                else if (col.BestFit)
                {
                    sw.Write(" bestFit=\"1\"");
                }
                sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", col.Width));

                if (col.OutlineLevel > 0)
                {
                    sw.Write($" outlineLevel=\"{col.OutlineLevel}\" ");
                    if (col.Collapsed)
                    {
                        sw.Write(" collapsed=\"1\"");
                    }
                }
                if (col.Phonetic)
                {
                    sw.Write(" phonetic=\"1\"");
                }

                var styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
                if (styleID > 0)
                {
                    sw.Write($" style=\"{styleID}\"");
                }
                sw.Write("/>");
            }
            if (!first)
            {
                sw.Write($"</{prefix}cols>");
            }
        }

        /// <summary>
        /// Check all Shared formulas that the first cell has not been deleted.
        /// If so create a standard formula of all cells in the formula .
        /// </summary>
        private void FixSharedFormulas()
        {
            var remove = new List<int>();
            foreach (var f in _ws._sharedFormulas.Values)
            {
                var addr = new ExcelAddressBase(f.Address);
                var shIx = _ws._formulas.GetValue(addr._fromRow, addr._fromCol);
                if (!(shIx is int) || (shIx is int && (int)shIx != f.Index))
                {
                    for (var row = addr._fromRow; row <= addr._toRow; row++)
                    {
                        for (var col = addr._fromCol; col <= addr._toCol; col++)
                        {
                            if (!(addr._fromRow == row && addr._fromCol == col))
                            {
                                var fIx = _ws._formulas.GetValue(row, col);
                                if (fIx is int && (int)fIx == f.Index)
                                {
                                    _ws._formulas.SetValue(row, col, f.GetFormula(row, col, _ws.Name));
                                }
                            }
                        }
                    }
                    remove.Add(f.Index);
                }
            }
            remove.ForEach(i => _ws._sharedFormulas.Remove(i));
        }

        // get StyleID without cell style for UpdateRowCellData
        internal int GetStyleIdDefaultWithMemo(int row, int col)
        {
            int v = 0;
            if (_ws.ExistsStyleInner(row, 0, ref v)) //First Row
            {
                return v;
            }
            else // then column
            {
                if (!columnStyles.ContainsKey(col))
                {
                    if (_ws.ExistsStyleInner(0, col, ref v))
                    {
                        columnStyles.Add(col, v);
                    }
                    else
                    {
                        int r = 0, c = col;
                        if (_ws._values.PrevCell(ref r, ref c))
                        {
                            var val = _ws._values.GetValue(0, c);
                            var column = (ExcelColumn)(val._value);
                            if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                            {
                                columnStyles.Add(col, val._styleId);
                            }
                            else
                            {
                                columnStyles.Add(col, 0);
                            }
                        }
                        else
                        {
                            columnStyles.Add(col, 0);
                        }
                    }
                }
                return columnStyles[col];
            }
        }

        private object GetFormulaValue(object v, string prefix)
        {
            if (v != null && v.ToString() != "")
            {
                return $"<{prefix}v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(v, _ws.Workbook.Date1904))}</{prefix}v>";
            }
            else
            {
                return "";
            }
        }

        private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row, string prefix)
        {
            if (prevRow != -1) cache.Append($"</{prefix}row>");
            //ulong rowID = ExcelRow.GetRowID(SheetID, row);
            cache.Append($"<{prefix}row r=\"{row}\"");
            RowInternal currRow = _ws.GetValueInner(row, 0) as RowInternal;
            if (currRow != null)
            {

                // if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
                if (currRow.Hidden == true)
                {
                    cache.Append(" hidden=\"1\"");
                }
                if (currRow.Height >= 0)
                {
                    cache.AppendFormat(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
                    if (currRow.CustomHeight)
                    {
                        cache.Append(" customHeight=\"1\"");
                    }
                }

                if (currRow.OutlineLevel > 0)
                {
                    cache.AppendFormat(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
                    if (currRow.Collapsed)
                    {
                        cache.Append(" collapsed=\"1\"");
                    }
                }
                if (currRow.Phonetic)
                {
                    cache.Append(" ph=\"1\"");
                }
            }
            var s = _ws.GetStyleInner(row, 0);
            if (s > 0)
            {
                cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID < 0 ? 0 : cellXfs[s].newID);
            }
            cache.Append(">");
        }

        private string GetDataTableAttributes(Formulas f)
        {
            var attributes = " ";
            if (f.IsDataTableRow)
            {
                attributes += "dtr=\"1\" ";
            }
            if (f.DataTableIsTwoDimesional)
            {
                attributes += "dt2D=\"1\" ";
            }
            if (f.FirstCellDeleted)
            {
                attributes += "del1=\"1\" ";
            }
            if (f.SecondCellDeleted)
            {
                attributes += "del2=\"1\" ";
            }
            if (string.IsNullOrEmpty(f.R1CellAddress) == false)
            {
                attributes += $"r1=\"{f.R1CellAddress}\" ";
            }
            if (string.IsNullOrEmpty(f.R2CellAddress) == false)
            {
                attributes += $"r2=\"{f.R2CellAddress}\" ";
            }
            return attributes;
        }

        /// <summary>
        /// Insert row and cells into the XML document
        /// </summary>
        private void UpdateRowCellData(StreamWriter sw, string prefix)
        {
            ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;

            int row = -1;
            string mdAttr = "";
            string mdAttrForFTag = "";
            var sheetDataTag = prefix + "sheetData";
            var cTag = prefix + "c";
            var fTag = prefix + "f";
            var vTag = prefix + "v";

            StringBuilder sbXml = new StringBuilder();
            var ss = _package.Workbook._sharedStrings;
            var cache = new StringBuilder();
            cache.Append($"<{sheetDataTag}>");


            FixSharedFormulas(); //Fixes Issue #32

            var hasMd = _ws._metadataStore.HasValues;
            columnStyles = new Dictionary<int, int>();
            var cse = new CellStoreEnumerator<ExcelValue>(_ws._values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Column > 0)
                {
                    var val = cse.Value;
                    int styleID = cellXfs[(val._styleId == 0 ? GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId)].newID;
                    styleID = styleID < 0 ? 0 : styleID;
                    //Add the row element if it's a new row
                    if (cse.Row != row)
                    {
                        WriteRow(cache, cellXfs, row, cse.Row, prefix);
                        row = cse.Row;
                    }
                    object v = val._value;
                    object formula = _ws._formulas.GetValue(cse.Row, cse.Column);
                    if (hasMd)
                    {
                        mdAttr = "";
                        if (_ws._metadataStore.Exists(cse.Row, cse.Column))
                        {
                            MetaDataReference md = _ws._metadataStore.GetValue(cse.Row, cse.Column);
                            if (md.cm > 0)
                            {
                                mdAttr = $" cm=\"{md.cm}\"";
                            }
                            if (md.vm > 0)
                            {
                                mdAttr += $" vm=\"{md.vm}\"";
                            }
                        }
                    }
                    if (formula is int sfId)
                    {
                        if (!_ws._sharedFormulas.ContainsKey(sfId))
                        {
                            throw (new InvalidDataException($"SharedFormulaId {sfId} not found on Worksheet {_ws.Name} cell {cse.CellAddress}, SharedFormulas Count {_ws._sharedFormulas.Count}"));
                        }
                        var f = _ws._sharedFormulas[sfId];

                        //Set calc attributes for array formula. We preserve them from load only at this point.
                        if (hasMd)
                        {
                            mdAttrForFTag = "";
                            if (_ws._metadataStore.Exists(cse.Row, cse.Column))
                            {
                                MetaDataReference md = _ws._metadataStore.GetValue(cse.Row, cse.Column);
                                if (md.aca)
                                {
                                    mdAttrForFTag = $" aca=\"1\"";
                                }
                                if (md.ca)
                                {
                                    mdAttrForFTag += $" ca=\"1\"";
                                }
                            }
                        }

                        if (f.Address.IndexOf(':') > 0)
                        {
                            if (f.StartCol == cse.Column && f.StartRow == cse.Row)
                            {
                                if (f.FormulaType == FormulaType.Array)
                                {
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"array\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                                }
                                else if (f.FormulaType == FormulaType.DataTable)
                                {
                                    var dataTableAttributes = GetDataTableAttributes(f);
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"dataTable\"{dataTableAttributes} {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}></{cTag}>");
                                }
                                else
                                {
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"shared\" si=\"{sfId}\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                                }

                            }
                            else if (f.FormulaType == FormulaType.Array)
                            {
                                string fElement;
                                if (string.IsNullOrEmpty(mdAttrForFTag) == false)
                                {
                                    fElement = $"<{fTag} {mdAttrForFTag}/>";
                                }
                                else
                                {
                                    fElement = $"";
                                }
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>{fElement}{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else if (f.FormulaType == FormulaType.DataTable)
                            {
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else
                            {
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><f t=\"shared\" si=\"{sfId}\" {mdAttrForFTag}/>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                        }
                        else
                        {
                            // We can also have a single cell array formula
                            if (f.FormulaType == FormulaType.Array)
                            {
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{string.Format("{0}:{1}", f.Address, f.Address)}\" t=\"array\"{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else
                            {
                                cache.Append($"<{cTag} r=\"{f.Address}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                                cache.Append($"<{fTag}{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                        }
                    }
                    else if (formula != null && formula.ToString() != "")
                    {
                        cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                        cache.Append($"<{fTag}>{ConvertUtil.ExcelEscapeAndEncodeString(formula.ToString())}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                    }
                    else
                    {
                        if (v == null && styleID > 0)
                        {
                            cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{mdAttr}/>");
                        }
                        else if (v != null)
                        {
                            if (v is System.Collections.IEnumerable enumResult && !(v is string))
                            {
                                var e = enumResult.GetEnumerator();
                                if (e.MoveNext() && e.Current != null)
                                    v = e.Current;
                                else
                                    v = string.Empty;
                            }
                            if ((TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan) && !(v is char))
                            {
                                //string sv = GetValueForXml(v);
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v)}{mdAttr}>");
                                cache.Append($"{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else
                            {
                                var s = Convert.ToString(v);
                                if (s == null) //If for example a struct 
                                {
                                    s = v.ToString();
                                    if (s == null)
                                    {
                                        s = "";
                                    }
                                }
                                int ix;
                                if (!ss.ContainsKey(s))
                                {
                                    ix = ss.Count;
                                    ss.Add(s, new ExcelWorkbook.SharedStringItem() { isRichText = _ws._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.RichText), pos = ix });
                                }
                                else
                                {
                                    ix = ss[s].pos;
                                }
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\" t=\"s\"{mdAttr}>");
                                cache.Append($"<{vTag}>{ix}</{vTag}></{cTag}>");
                            }
                        }
                    }
                }
                else  //ExcelRow
                {
                    WriteRow(cache, cellXfs, row, cse.Row, prefix);
                    row = cse.Row;
                }
                if (cache.Length > 0x600000)
                {
                    sw.Write(cache.ToString());
                    sw.Flush();
                    cache.Length = 0;
                }
            }
            columnStyles = null;

            if (row != -1) cache.Append($"</{prefix}row>");
            cache.Append($"</{prefix}sheetData>");
            sw.Write(cache.ToString());
            sw.Flush();
        }

        /// <summary>
        /// Update merged cells
        /// </summary>
        /// <param name="sw">The writer</param>
        /// <param name="prefix">Namespace prefix for the main schema</param>
        private void UpdateMergedCells(StreamWriter sw, string prefix)
        {
            sw.Write($"<{prefix}mergeCells>");
            foreach (string address in _ws._mergedCells.Distinct())
            {
                sw.Write($"<{prefix}mergeCell ref=\"{address}\" />");
            }
            sw.Write($"</{prefix}mergeCells>");
        }

        private void WriteDataValidationAttributes(ref StringBuilder cache, int i)
        {
            if (_ws.DataValidations[i].ValidationType != null)
            {
                cache.Append($"type=\"{_ws.DataValidations[i].ValidationType.TypeToXmlString()}\" ");
            }

            if (_ws.DataValidations[i].ErrorStyle != ExcelDataValidationWarningStyle.undefined)
            {
                cache.Append($"errorStyle=\"{_ws.DataValidations[i].ErrorStyle.ToString()}\" ");
            }

            //support IME mode here?

            if (_ws.DataValidations[i].Operator != 0)
            {
                cache.Append($"operator=\"{_ws.DataValidations[i].Operator.ToString()}\" ");
            }


            //Note that if false excel does not write these properties out so we don't either.
            if (_ws.DataValidations[i].AllowBlank == true)
            {
                cache.Append($"allowBlank=\"1\" ");
            }

            if (_ws.DataValidations[i] is ExcelDataValidationList)
            {
                if ((_ws.DataValidations[i] as ExcelDataValidationList).HideDropDown == true)
                {
                    cache.Append($"showDropDown=\"1\" ");
                }
            }

            if (_ws.DataValidations[i].ShowInputMessage == true)
            {
                cache.Append($"showInputMessage=\"1\" ");
            }

            if (_ws.DataValidations[i].ShowErrorMessage == true)
            {
                cache.Append($"showErrorMessage=\"1\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].ErrorTitle) == false)
            {
                cache.Append($"errorTitle=\"{_ws.DataValidations[i].ErrorTitle}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].Error) == false)
            {
                cache.Append($"error=\"{_ws.DataValidations[i].Error}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].PromptTitle) == false)
            {
                cache.Append($"promptTitle=\"{_ws.DataValidations[i].PromptTitle}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].Prompt) == false)
            {
                cache.Append($"prompt=\"{_ws.DataValidations[i].Prompt}\" ");
            }

            if (_ws.DataValidations.HasValidationType(InternalValidationType.DataValidation))
            {
                cache.Append($"sqref=\"{_ws.DataValidations[i].Address}\" ");
            }

            cache.Append($"xr:uid=\"{_ws.DataValidations[i].Uid}\"");

            cache.Append(">");
        }

        private void WriteDataValidation(ref StringBuilder cache, string prefix, int i, string extNode = "")
        {
            cache.Append($"<{prefix}dataValidation ");
            WriteDataValidationAttributes(ref cache, i);

            if (_ws.DataValidations[i].ValidationType.Type != eDataValidationType.Any)
            {
                string endExtNode = "";
                if (extNode != "")
                {
                    endExtNode = $"</{extNode}>";
                    extNode = $"<{extNode}>";
                }

                switch (_ws.DataValidations[i].ValidationType.Type)
                {
                    case eDataValidationType.TextLength:
                    case eDataValidationType.Whole:
                        var intType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>;
                        string intString = ((ExcelDataValidationFormulaInt)intType.Formula).GetXmlValue();
                        string intString2 = ((ExcelDataValidationFormulaInt)intType.Formula2).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{intString}{endExtNode}</{prefix}formula1>");
                        cache.Append($"<{prefix}formula2>{extNode}{intString2}{endExtNode}</{prefix}formula2>");
                        break;
                    case eDataValidationType.Decimal:
                        var decimalType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>;
                        string decimalString = ((ExcelDataValidationFormulaDecimal)decimalType.Formula).GetXmlValue();
                        string decimalString2 = ((ExcelDataValidationFormulaDecimal)decimalType.Formula2).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{decimalString}{endExtNode}</{prefix}formula1>");
                        cache.Append($"<{prefix}formula2>{extNode}{decimalString2}{endExtNode}</{prefix}formula2>");
                        break;
                    case eDataValidationType.List:
                        var listType = _ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>;
                        string listString = ((ExcelDataValidationFormulaList)listType.Formula).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{listString}{endExtNode}</{prefix}formula1>");
                        break;
                    case eDataValidationType.Time:
                        var timeType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>;
                        string timeString = ((ExcelDataValidationFormulaTime)timeType.Formula).GetXmlValue();
                        string timeString2 = ((ExcelDataValidationFormulaTime)timeType.Formula2).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{timeString}{endExtNode}</{prefix}formula1>");
                        cache.Append($"<{prefix}formula2>{extNode}{timeString2}{endExtNode}</{prefix}formula2>");
                        break;
                    case eDataValidationType.DateTime:
                        var dateTimeType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>;
                        string dateTimeString = ((ExcelDataValidationFormulaDateTime)dateTimeType.Formula).GetXmlValue();
                        string dateTimeString2 = ((ExcelDataValidationFormulaDateTime)dateTimeType.Formula2).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{dateTimeString}{endExtNode}</{prefix}formula1>");
                        cache.Append($"<{prefix}formula2>{extNode}{dateTimeString2}{endExtNode}</{prefix}formula2>");
                        break;
                    case eDataValidationType.Custom:
                        var customType = _ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormula>;
                        string customString = ((ExcelDataValidationFormulaCustom)customType.Formula).GetXmlValue();

                        cache.Append($"<{prefix}formula1>{extNode}{customString}{endExtNode}</{prefix}formula1>");
                        break;
                    default:
                        throw new Exception("UNKNOWN TYPE IN WriteDataValidation");
                }

                if (extNode != "")
                {
                    cache.Append($"<xm:sqref>{_ws.DataValidations[i].Address}</xm:sqref>");
                }
            }

            //write adress if extLst
            cache.Append($"</{prefix}dataValidation>");
        }

        private void UpdateDataValidation(StreamWriter sw, string prefix, string extraAttribute = "")
        {
            var cache = new StringBuilder();
            InternalValidationType type;
            string extNode = "";

            if (extraAttribute == "")
            {

                cache.Append($"<{prefix}dataValidations count=\"{_ws.DataValidations.GetNonExtLstCount()}\">");
                type = InternalValidationType.DataValidation;
            }
            else
            {
                cache.Append($"<{prefix}dataValidations {extraAttribute} count=\"{_ws.DataValidations.GetExtLstCount()}\">");
                type = InternalValidationType.ExtLst;
                extNode = "xm:f";
            }

            for (int i = 0; i < _ws.DataValidations.Count; i++)
            {
                if (_ws.DataValidations[i].InternalValidationType == type)
                {
                    WriteDataValidation(ref cache, prefix, i, extNode);
                }
            }

            cache.Append($"</{prefix}dataValidations>");

            sw.Write(cache.ToString());
            sw.Flush();
        }
    }
}
