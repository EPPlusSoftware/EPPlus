/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2023       EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.ExcelXMLWriter;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using static OfficeOpenXml.ExcelWorkbook;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.Core.Worksheet.XmlWriter
{
    internal class WorksheetXmlWriter
    {
        ExcelWorksheet _ws;
        ExcelPackage _package;
        private Dictionary<int, int> columnStyles = null;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="package"></param>
        public WorksheetXmlWriter(ExcelWorksheet worksheet, ExcelPackage package)
        {
            _ws = worksheet;
            _package = package;
        }

        /// <summary>
        /// Replaces placeholder nodes by writing the system's held information
        /// </summary>
        /// <param name="sw">The streamwriter file info is written to</param>
        /// <param name="xml">The original XML</param>
        /// <param name="startOfNode">Start position of the current node</param>
        /// <param name="endOfNode">End position of the current node</param>
        internal void WriteNodes(StreamWriter sw, string xml, ref int startOfNode, ref int endOfNode)
        {
            var prefix = _ws.GetNameSpacePrefix();

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

            FindNodePositionAndClearIt(sw, xml, "conditionalFormatting", ref startOfNode, ref endOfNode);
            sw.Write(UpdateConditionalFormattings(prefix));

            if (_ws.GetNode("d:dataValidations") != null)
            {
                FindNodePositionAndClearIt(sw, xml, "dataValidations", ref startOfNode, ref endOfNode);
                if (_ws.DataValidations.GetNonExtLstCount() > 0)
                {
                    sw.Write(UpdateDataValidation(prefix));
                }
            }

            FindNodePositionAndClearIt(sw, xml, "hyperlinks", ref startOfNode, ref endOfNode);
            UpdateHyperLinks(sw, prefix);

            FindNodePositionAndClearIt(sw, xml, "rowBreaks", ref startOfNode, ref endOfNode);
            UpdateRowBreaks(sw, prefix);

            FindNodePositionAndClearIt(sw, xml, "colBreaks", ref startOfNode, ref endOfNode);
            UpdateColBreaks(sw, prefix);

            ExtLstHelper extLst = new ExtLstHelper(xml, _ws);

            FindNodePositionAndClearIt(sw, xml, "extLst", ref startOfNode, ref endOfNode);

            //Careful. Ensure that we only do appropriate extLst things when there are objects to operate on.
            //Creating an empty DataValidations Node in ExtLst for example generates a corrupt excelfile that passes validation tool checks.
            if (_ws.DataValidations.GetExtLstCount() != 0)
            {
                extLst.InsertExt(ExtLstUris.DataValidationsUri, UpdateExtLstDataValidations(prefix), "");
            }

            extLst.InsertExt(ExtLstUris.ConditionalFormattingUri, UpdateExtLstConditionalFormatting(), "");

            if (extLst.extCount != 0)
            {
                sw.Write(extLst.GetWholeExtLst());
            }

            sw.Write(xml.Substring(endOfNode, xml.Length - endOfNode));
        }

        internal void FindNodePositionAndClearItInit(StreamWriter sw, string xml, string nodeName,
            ref int start, ref int end)
        {
            start = end;
            GetBlock.Pos(xml, nodeName, ref start, ref end);

            sw.Write(xml.Substring(0, start));
        }

        internal void FindNodePositionAndClearIt(StreamWriter sw, string xml, string nodeName,
            ref int start, ref int end)
        {
            int oldEnd = end;
            GetBlock.Pos(xml, nodeName, ref start, ref end);

            sw.Write(xml.Substring(oldEnd, start - oldEnd));
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
                var col = cse.Value._value as ExcelColumn;
                if (col == null) continue;
                if (first)
                {
                    sw.Write($"<{prefix}cols>");
                    first = false;
                }
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
                            var column = (ExcelColumn)val._value;
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

        private string GetDataTableAttributes(SharedFormula f)
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
            var nsf = _package.Workbook.FormulaParser.ParsingContext.Configuration.FunctionRepository.NamespaceFunctions;
            StringBuilder sbXml = new StringBuilder();
            var ssLookup = _package.Workbook._sharedStringsLookup;
            var ssList = _package.Workbook._sharedStringsListNew;
            var cache = new StringBuilder();
            cache.Append($"<{sheetDataTag}>");

            FixSharedFormulas(); //Fixes Issue #32

            var hasMd = _ws._metadataStore.HasValues || _ws.Workbook.HasMetadataPart;
            var hasFlags = _ws._flags.HasValues;
            var hasRd = false;
            columnStyles = new Dictionary<int, int>();
            var cse = new CellStoreEnumerator<ExcelValue>(_ws._values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Column > 0)
                {
                    var val = cse.Value;
                    int styleID = cellXfs[val._styleId == 0 ? GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId].newID;
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
                        if (v is ExcelErrorValue error)
                        {
                            if (error.Type == eErrorType.Spill || error.Type == eErrorType.Calc)
                            {
                                v = ErrorValues.ValueError;
                                SetMetaDataForError(cse, error);
                                hasRd = true;
                            }
                        }

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
                            throw new InvalidDataException($"SharedFormulaId {sfId} not found on Worksheet {_ws.Name} cell {cse.CellAddress}, SharedFormulas Count {_ws._sharedFormulas.Count}");
                        }
                        var f = _ws._sharedFormulas[sfId];

                        //Set calc attributes for array formula. We preserve them from load only at this point.
                        if (hasFlags)
                        {
                            mdAttrForFTag = "";
                            if (_ws._flags.Exists(cse.Row, cse.Column))
                            {                                
                                if (_ws._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.CellFlagAlwaysCalculateArray))
                                {
                                    mdAttrForFTag = $" aca=\"1\"";
                                }
                                if (_ws._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.CellFlagCalculateCell))
                                {
                                    mdAttrForFTag += $" ca=\"1\"";
                                }
                            }
                        }
                        if(f._hasUpdatedNamespace==false) f.UpdateFormulaNamespaces(nsf);
                        if (f.Address.IndexOf(':') > 0)
                        {
                            if (f.StartCol == cse.Column && f.StartRow == cse.Row)
                            {
                                if (f.FormulaType == FormulaType.Array)
                                {
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"array\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula, false)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                                }
                                else if (f.FormulaType == FormulaType.DataTable)
                                {
                                    var dataTableAttributes = GetDataTableAttributes(f);
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"dataTable\"{dataTableAttributes} {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula, false)}</{fTag}></{cTag}>");
                                }
                                else
                                {
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"shared\" si=\"{sfId}\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula, false)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
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
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{string.Format("{0}:{1}", f.Address, f.Address)}\" t=\"array\"{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula, false)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else
                            {
                                cache.Append($"<{cTag} r=\"{f.Address}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                                cache.Append($"<{fTag}{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula, false)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                        }
                    }
                    else if (formula != null && formula.ToString() != "")
                    {
                        var f= SharedFormula.UpdateFormulaNamespaces(formula.ToString(), nsf);
                        cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                        cache.Append($"<{fTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f, false)}</{fTag}>{GetFormulaValue(v, prefix)}</{cTag}>");
                    }
                    else
                    {
                        if (v == null && styleID > 0)
                        {
                            cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{mdAttr}/>");
                        }
                        else if (v != null)
                        {
                            if (v is ExcelRichTextCollection rt)
                            {
                                var s = rt.GetXML();
                                if (!ssLookup.TryGetValue(s, out int ix))
                                {
                                    ix = ssLookup.Count;
                                    ssLookup.Add(s, ix);
                                    ssList.Add(new SharedStringRichTextItem() { RichText = rt, Position = ix });
                                }
                                cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\" t=\"s\"{mdAttr}>");
                                cache.Append($"<{vTag}>{ix}</{vTag}></{cTag}>");
                            }
                            else
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
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v)}{mdAttr}>");
                                    cache.Append($"{GetFormulaValue(v, prefix)}</{cTag}>");
                                }
                                else if (v is ExcelErrorValue e)
                                {
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
                                    if (!ssLookup.TryGetValue(s, out ix))
                                    {
                                        ix = ssLookup.Count;
                                        ssLookup.Add(s, ix);
                                        ssList.Add(new SharedStringTextItem() { Text = s, Position = ix });
                                    }
                                    cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\" t=\"s\"{mdAttr}>");
                                    cache.Append($"<{vTag}>{ix}</{vTag}></{cTag}>");
                                }
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

            if(hasRd)
            {
                _ws.Workbook.RichData.SetHasValuesOnParts();
            }
        }

        private void SetMetaDataForError(CellStoreEnumerator<ExcelValue> cse, ExcelErrorValue error)
        {
            var richData = _package.Workbook.RichData;
            var metadata = _package.Workbook.Metadata;
            var md = _ws._metadataStore.GetValue(cse.Row, cse.Column);
            if(md.vm >= 0 && IsMdSameError(metadata, md, error, cse.Row, cse.Column))
            {
                return;
            }
            switch (error.Type)
            {
                case eErrorType.Spill:
                    var spillError = (ExcelRichDataErrorValue)error;
                    if(spillError.IsPropagated)
                    {
                        richData.Values.AddPropagated(eErrorType.Spill);
                    }
                    else
                    {
                        richData.Values.AddErrorSpill(spillError);                    
                    }
                    break;
                case eErrorType.Calc:
                    richData.Values.AddError(eErrorType.Calc, "1");
                    break;
                default:
                    return;
            }
            var fmdRichDataCollection = metadata.GetFutureMetadataRichDataCollection();
            var rdItem = new ExcelFutureMetadataRichData(richData.Values.Items.Count-1);
            fmdRichDataCollection.Types.Add(rdItem);
            var mdItem = new ExcelMetadataItem();
            mdItem.Records.Add(new ExcelMetadataRecord(metadata.RichDataTypeIndex, fmdRichDataCollection.Types.Count - 1));
            metadata.ValueMetadata.Add(mdItem);

            md.vm = metadata.ValueMetadata.Count;
            _ws._metadataStore.SetValue(cse.Row, cse.Column, md);
        }

        private bool IsMdSameError(ExcelMetadata metadata, MetaDataReference md, ExcelErrorValue error, int row, int column)
        {
            if (md.vm==0 || md.vm>=metadata.ValueMetadata.Count) return false;
            var vm = metadata.ValueMetadata[md.vm-1];
            if (vm.Records.Count > 0 && vm.Records[0].ValueTypeIndex >=0)
            {
                var richData = _package.Workbook.RichData;
                if (richData.Values.Items.Count > vm.Records[0].ValueTypeIndex)
                {
                    var rd = richData.Values.Items[vm.Records[0].ValueTypeIndex];
                    if (rd.Structure.Type.Equals("_error"))
                    {
                        switch (error.Type)
                        {
                            case eErrorType.Calc:
                                if (rd.Values[0] == "13")
                                {
                                    return true;
                                }
                                break;
                            case eErrorType.Spill:
                                var rdError = (ExcelRichDataErrorValue)error; 
                                if (rd.HasValue(["errorType", "colOffset", "rwOffset"], ["8", rdError.SpillColOffset.ToString(CultureInfo.InvariantCulture), rdError.SpillColOffset.ToString(CultureInfo.InvariantCulture)]))
                                {
                                    return true;
                                }
                                break;
                        }
                    }
                }
            }
            return false;
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
            if (_ws.DataValidations[i].ValidationType != null &&
                _ws.DataValidations[i].ValidationType.Type != eDataValidationType.Any)
            {
                cache.Append($"type=\"{_ws.DataValidations[i].ValidationType.TypeToXmlString()}\" ");
            }

            if (_ws.DataValidations[i].ErrorStyle != ExcelDataValidationWarningStyle.undefined)
            {
                cache.Append($"errorStyle=\"{_ws.DataValidations[i].ErrorStyle.ToEnumString()}\" ");
            }

            if (_ws.DataValidations[i].ImeMode != ExcelDataValidationImeMode.NoControl)
            {
                cache.Append($"imeMode=\"{_ws.DataValidations[i].ImeMode.ToEnumString()}\" ");
            }

            if (_ws.DataValidations[i].Operator != 0)
            {
                cache.Append($"operator=\"{_ws.DataValidations[i].Operator.ToEnumString()}\" ");
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
                cache.Append($"errorTitle=\"{_ws.DataValidations[i].ErrorTitle.EncodeXMLAttribute()}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].Error) == false)
            {
                cache.Append($"error=\"{_ws.DataValidations[i].Error.EncodeXMLAttribute()}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].PromptTitle) == false)
            {
                cache.Append($"promptTitle=\"{_ws.DataValidations[i].PromptTitle.EncodeXMLAttribute()}\" ");
            }

            if (string.IsNullOrEmpty(_ws.DataValidations[i].Prompt) == false)
            {
                cache.Append($"prompt=\"{_ws.DataValidations[i].Prompt.EncodeXMLAttribute()}\" ");
            }

            if (_ws.DataValidations[i].InternalValidationType == InternalValidationType.DataValidation)
            {
                //Might need encode xml
                cache.Append($"sqref=\"{_ws.DataValidations[i].Address.ToString().Replace(",", " ")}\" ");
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
                        WriteDataValidationFormulas(intType.Formula, intType.Formula2, cache,
                            prefix, extNode, endExtNode, _ws.DataValidations[i].Operator);
                        break;
                    case eDataValidationType.Decimal:
                        var decimalType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>;
                        WriteDataValidationFormulas(decimalType.Formula, decimalType.Formula2, cache,
                            prefix, extNode, endExtNode, _ws.DataValidations[i].Operator);
                        break;
                    case eDataValidationType.List:
                        var listType = _ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>;
                        WriteDataValidationFormulaSingle(listType.Formula, cache, prefix, extNode, endExtNode);
                        break;
                    case eDataValidationType.Time:
                        var timeType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>;
                        WriteDataValidationFormulas(timeType.Formula, timeType.Formula2, cache,
                            prefix, extNode, endExtNode, _ws.DataValidations[i].Operator);
                        break;
                    case eDataValidationType.DateTime:
                        var dateTimeType = _ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>;
                        WriteDataValidationFormulas(dateTimeType.Formula, dateTimeType.Formula2, cache,
                            prefix, extNode, endExtNode, _ws.DataValidations[i].Operator);
                        break;
                    case eDataValidationType.Custom:
                        var customType = _ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormula>;
                        WriteDataValidationFormulaSingle(customType.Formula, cache, prefix, extNode, endExtNode);
                        break;
                    default:
                        throw new Exception("UNKNOWN TYPE IN WriteDataValidation");
                }

                if (extNode != "")
                {
                    //write adress if extLst
                    cache.Append($"<xm:sqref>{_ws.DataValidations[i].Address.ToString().Replace(",", " ")}</xm:sqref>");
                }
            }

            cache.Append($"</{prefix}dataValidation>");
        }

        void WriteDataValidationFormulaSingle(IExcelDataValidationFormula formula,
            in StringBuilder cache, string prefix, string extNode, string endExtNode)
        {
            string string1 = ((ExcelDataValidationFormula)formula).GetXmlValue();

            if (string.IsNullOrEmpty(string1))
            {
                return;
            }

            string1 = ConvertUtil.ExcelEscapeAndEncodeString(string1);

            cache.Append($"<{prefix}formula1>{extNode}{string1}{endExtNode}</{prefix}formula1>");
        }

        void WriteDataValidationFormulas(IExcelDataValidationFormula formula1, IExcelDataValidationFormula formula2,
            in StringBuilder cache, string prefix, string extNode, string endExtNode, ExcelDataValidationOperator dvOperator)
        {
            string string1 = ((ExcelDataValidationFormula)formula1).GetXmlValue();
            string string2 = ((ExcelDataValidationFormula)formula2).GetXmlValue();

            if (string.IsNullOrEmpty(string1) && string.IsNullOrEmpty(string2))
            {
                return;
            }

            //Note that formula1 must be written even when string1 is empty
            string1 = ConvertUtil.ExcelEscapeAndEncodeString(string1);
            cache.Append($"<{prefix}formula1>{extNode}{string1}{endExtNode}</{prefix}formula1>");

            if (!string.IsNullOrEmpty(string2) &&
                (dvOperator == ExcelDataValidationOperator.between || dvOperator == ExcelDataValidationOperator.notBetween))
            {
                string2 = ConvertUtil.ExcelEscapeAndEncodeString(string2);
                cache.Append($"<{prefix}formula2>{extNode}{string2}{endExtNode}</{prefix}formula2>");
            }
        }
        private StringBuilder UpdateDataValidation(string prefix, string extraAttribute = "")
        {
            var cache = new StringBuilder();
            InternalValidationType type;
            string extNode = "";

            if (extraAttribute == "")
            {
                if (_ws.DataValidations.GetNonExtLstCount() > 0)
                {
                    cache.Append($"<{prefix}dataValidations count=\"{_ws.DataValidations.GetNonExtLstCount()}\">");
                    type = InternalValidationType.DataValidation;
                }
                else
                {
                    return cache;
                }
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

            return cache;
        }

        /// <summary>
        /// Update xml with hyperlinks 
        /// </summary>
        /// <param name="sw">The stream</param>
        /// <param name="prefix">The namespace prefix for the main schema</param>
        private void UpdateHyperLinks(StreamWriter sw, string prefix)
        {
			ClearHyperlinks();
			Dictionary<string, string> hyps = new Dictionary<string, string>();
            var cse = new CellStoreEnumerator<Uri>(_ws._hyperLinks);
            bool first = true;
            while (cse.Next())
            {
                var uri = _ws._hyperLinks.GetValue(cse.Row, cse.Column);
                if (first && uri != null)
                {
                    sw.Write($"<{prefix}hyperlinks>");
                    first = false;
                }
                var hl = uri as ExcelHyperLink;
                if (hl != null && !string.IsNullOrEmpty(hl.ReferenceAddress) && uri.OriginalString.StartsWith("xl://internal"))
                {
                    var address = _ws.Cells[cse.Row, cse.Column, cse.Row + hl.RowSpan, cse.Column + hl.ColSpan].Address;
                    var location = ExcelCellBase.GetFullAddress(SecurityElement.Escape(_ws.Name), SecurityElement.Escape(hl.ReferenceAddress));
                    var display = string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"";
                    var tooltip = string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"";
                    sw.Write($"<{prefix}hyperlink ref=\"{address}\" location=\"{location}\"{display}{tooltip}/>");
                }
                else if (uri != null)
                {
                    string id;
                    Uri hyp;
                    string target = ""; 
                    if (hl != null)
                    {
                        if (hl.Target != null && hl.OriginalString.StartsWith("Invalid:Uri", StringComparison.OrdinalIgnoreCase))
                        {
                            target = hl.Target;
                        }
                        hyp = hl.OriginalUri;
                    }
                    else
                    {
                        hyp = uri;
                    }

                    if (hyps.ContainsKey(hyp.OriginalString) && string.IsNullOrEmpty(target))
                    {
                        id = hyps[hyp.OriginalString];
                    }
                    else
                    {
                        ZipPackageRelationship relationship;
                        if (string.IsNullOrEmpty(target))
                        {
                            relationship = _ws.Part.CreateRelationship(hyp, TargetMode.External, ExcelPackage.schemaHyperlink);
                        }
                        else
                        {
                            relationship = _ws.Part.CreateRelationship(target, TargetMode.External, ExcelPackage.schemaHyperlink);
                        }
                        if (hl != null)
                        {
                            var display = string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"";
                            var toolTip = string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"";
                            var location = string.IsNullOrEmpty(hl.ReferenceAddress) ? "" : " location=\"" + SecurityElement.Escape(hl.ReferenceAddress) + "\"";
                            sw.Write($"<{prefix}hyperlink ref=\"{ExcelCellBase.GetAddress(cse.Row, cse.Column)}\"{location}{display}{toolTip} r:id=\"{relationship.Id}\"/>");
                        }
                        else
                        {
                            sw.Write($"<{prefix}hyperlink ref=\"{ExcelCellBase.GetAddress(cse.Row, cse.Column)}\" r:id=\"{relationship.Id}\"/>");
                        }
                    }
                }
            }
            if (!first)
            {
                sw.Write($"</{prefix}hyperlinks>");
            }
        }

		private void ClearHyperlinks()
		{
            var maxRelId = 0;
			foreach(var rel in _ws.Part._rels.ToList())
            {
                if(rel.RelationshipType.Equals(ExcelPackage.schemaHyperlink, StringComparison.OrdinalIgnoreCase))
                {
                    _ws.Part.DeleteRelationship(rel.Id);
                }
                else
                {
					if (rel.Id.StartsWith("rId"))
					{
						if (int.TryParse(rel.Id.Substring(3), out int id))
						{
							if (id > maxRelId)
							{
								maxRelId = id;
							}
						}
					}

				}
			}
            _ws.Part.SetMaxRelId(maxRelId+1);
		}

		private void UpdateRowBreaks(StreamWriter sw, string prefix)
        {
            StringBuilder breaks = new StringBuilder();
            int count = 0;
            var cse = new CellStoreEnumerator<ExcelValue>(_ws._values, 0, 0, ExcelPackage.MaxRows, 0);
            while (cse.Next())
            {
                var row = cse.Value._value as RowInternal;
                if (row != null && row.PageBreak)
                {
                    breaks.AppendFormat($"<{prefix}brk id=\"{cse.Row}\" max=\"1048575\" man=\"1\"/>");
                    count++;
                }
            }
            if (count > 0)
            {
                sw.Write(string.Format($"<{prefix}rowBreaks count=\"{count}\" manualBreakCount=\"{count}\">{breaks.ToString()}</rowBreaks>"));
            }
        }

        private void UpdateColBreaks(StreamWriter sw, string prefix)
        {
            StringBuilder breaks = new StringBuilder();
            int count = 0;
            var cse = new CellStoreEnumerator<ExcelValue>(_ws._values, 0, 0, 0, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                var col = cse.Value._value as ExcelColumn;
                if (col != null && col.PageBreak)
                {
                    breaks.Append($"<{prefix}brk id=\"{cse.Column}\" max=\"16383\" man=\"1\"/>");
                    count++;
                }
            }
            if (count > 0)
            {
                sw.Write($"<colBreaks count=\"{count}\" manualBreakCount=\"{count}\">{breaks.ToString()}</colBreaks>");
            }
        }

        /// <summary>
        /// ExtLst updater for DataValidations
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        private string UpdateExtLstDataValidations(string prefix)
        {
            var cache = new StringBuilder();

            cache.Append($"<ext xmlns:x14=\"{ExcelPackage.schemaMainX14}\" uri=\"{ExtLstUris.DataValidationsUri}\">");

            prefix = "x14:";
            cache.Append
            (
            UpdateDataValidation(prefix, $"xmlns:xm=\"{ExcelPackage.schemaMainXm}\"")

            );
            cache.Append("</ext>");

            return cache.ToString();
        }

        private string UpdateConditionalFormattingAttributes(
           ExcelConditionalFormattingRule conditionalFormat)
        {
            StringBuilder cache = new StringBuilder();

            if (conditionalFormat.DxfId != -1)
            {
                cache.Append($"dxfId=\"{conditionalFormat.DxfId}\" ");
            }

            cache.Append($"priority=\"{conditionalFormat.Priority}\" ");

            if (conditionalFormat.StopIfTrue)
            {
                cache.Append($"stopIfTrue=\"1\" ");
            }

            if (conditionalFormat.AboveAverage == false)
            {
                cache.Append($"aboveAverage=\"0\" ");
            }

            if ((bool)conditionalFormat.Percent)
            {
                cache.Append($"percent=\"1\" ");
            }

            if ((bool)conditionalFormat.Bottom)
            {
                cache.Append($"bottom=\"1\" ");
            }

            if (conditionalFormat.Operator != null)
            {
                cache.Append($"operator=\"{conditionalFormat.Operator.ToEnumString()}\" ");
            }

            if (string.IsNullOrEmpty(conditionalFormat._text) == false)
            {
                cache.Append($"text=\"{conditionalFormat._text.EncodeXMLAttribute()}\" ");
            }

            if (conditionalFormat.TimePeriod != null)
            {
                cache.Append($"timePeriod=\"{conditionalFormat.TimePeriod.ToEnumString()}\" ");
            }

            if (conditionalFormat._rank != 0)
            {
                cache.Append($"rank=\"{conditionalFormat.Rank}\" ");
            }

            if (conditionalFormat.StdDev != 0)
            {
                cache.Append($"stdDev=\"{conditionalFormat.StdDev}\" ");
            }

            if ((bool)conditionalFormat.EqualAverage)
            {
                cache.Append("equalAverage=\"1\" ");
            }

            return cache.ToString();
        }

        private string UpdateExtLstConditionalFormatting()
        {
            var cache = new StringBuilder();

            var addressDict = new Dictionary<string, List<ExcelConditionalFormattingRule>>();

            //Find each extLst cfRule for given address
            foreach (var format in _ws.ConditionalFormatting)
            {
                if (format is ExcelConditionalFormattingRule cf)
                {
                    if (cf.IsExtLst)
                    {
                        var key = cf.Address?.Address ?? "";
                        if (addressDict.ContainsKey(key))
                        {
                            addressDict[key].Add(cf);
                        }
                        else
                        {
                            addressDict.Add(key, new List<ExcelConditionalFormattingRule>() { cf });
                        }
                    }
                }
            }

            if (addressDict.Count > 0)
            {
                string prefix = "x14:";
                cache.Append($"<ext xmlns:x14=\"{ExcelPackage.schemaMainX14}\" uri=\"{ExtLstUris.ConditionalFormattingUri}\">");

                cache.Append($"<{prefix}conditionalFormattings>");

                foreach (var ruleList in addressDict)
                {
                    for (int i = 0; i < ruleList.Value.Count; i++)
                    {
                        var format = ruleList.Value[i];

                        if (i == 0)
                        {
                            cache.Append($"<{prefix}conditionalFormatting xmlns:xm=\"{ExcelPackage.schemaMainXm}\"");
                            if (format.PivotTable)
                            {
                                cache.Append($" pivot=\"1\"");
                            }
                            cache.Append($">");
                        }

                        string uid;

                        if (format.Type == eExcelConditionalFormattingRuleType.DataBar)
                        {
                            var dataBar = (ExcelConditionalFormattingDataBar)format;
                            uid = dataBar.Uid.ToString();

                            cache.Append($"<{prefix}cfRule type=\"{format.Type.ToString().UnCapitalizeFirstLetter()}\" id=\"{uid}\">");

                            cache.Append($"<{prefix}dataBar minLength=\"{dataBar.LowValue.minLength}\" ");
                            cache.Append($"maxLength=\"{dataBar.HighValue.maxLength}\"");

                            if(dataBar.ShowValue == false)
                            {
                                cache.Append($" showValue=\"0\"");
                            }

                            if (dataBar.Border)
                            {
                                cache.Append($" border=\"1\"");
                            }

                            if(dataBar.Direction != eDatabarDirection.Context)
                            {
                                cache.Append($" direction=\"{dataBar.Direction.ToEnumString()}\"");
                            }

                            if (dataBar.Gradient == false)
                            {
                                cache.Append($" gradient=\"0\"");
                            }

                            if(dataBar.NegativeBarColorSameAsPositive)
                            {
                                cache.Append($" negativeBarColorSameAsPositive=\"1\"");
                            }

                            if (!dataBar.NegativeBarBorderColorSameAsPositive)
                            {
                                cache.Append($" negativeBarBorderColorSameAsPositive=\"0\"");
                            }

                            if(dataBar.AxisPosition != eExcelDatabarAxisPosition.Automatic)
                            {
                                cache.Append($" axisPosition=\"{dataBar.AxisPosition.ToEnumString()}\"");
                            }

                            cache.Append(">");

                            cache.Append($"<{prefix}cfvo type=\"{dataBar.LowValue.Type.ToString().UnCapitalizeFirstLetter()}\"");

                            if (dataBar.LowValue.HasValueOrFormula)
                            {
                                cache.Append(">");
                                if(!string.IsNullOrEmpty(dataBar.LowValue.Formula))
                                {
                                    cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(dataBar.LowValue.Formula)}</xm:f>");
                                }
                                else if(dataBar.LowValue.Value != double.NaN)
                                {
                                    cache.Append($"<xm:f>{dataBar.LowValue.Value}</xm:f>");
                                }
                                cache.Append($"</x14:cfvo>");
                            }
                            else
                            {
                                cache.Append("/>");
                            }

                            cache.Append($"<{prefix}cfvo type=\"{dataBar.HighValue.Type.ToString().UnCapitalizeFirstLetter()}\"");

                            if (dataBar.HighValue.HasValueOrFormula)
                            {
                                cache.Append(">");
                                if (!string.IsNullOrEmpty(dataBar.HighValue.Formula))
                                {
                                    cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(dataBar.HighValue.Formula)}</xm:f>");
                                }
                                else
                                {
                                    cache.Append($"<xm:f>{dataBar.HighValue.Value}</xm:f>");
                                }
                                cache.Append($"</x14:cfvo>");
                            }
                            else
                            {
                                cache.Append("/>");
                            }

                            if (dataBar.BorderColor != null && dataBar.Border == true)
                            {
                                WriteDxfColor(prefix, cache, dataBar.BorderColor, "borderColor");
                            }

                            if (dataBar.NegativeFillColor != null && dataBar.NegativeBarColorSameAsPositive == false)
                            {
                                WriteDxfColor(prefix, cache, dataBar.NegativeFillColor, "negativeFillColor");
                            }

                            if (dataBar.NegativeBorderColor != null && dataBar.NegativeBarColorSameAsPositive == false && dataBar.Border == true)
                            {
                                WriteDxfColor(prefix, cache, dataBar.NegativeBorderColor, "negativeBorderColor");
                            }

                            if (dataBar.AxisColor != null && dataBar.AxisPosition != eExcelDatabarAxisPosition.None)
                            {
                                WriteDxfColor(prefix, cache, dataBar.AxisColor, "axisColor");
                            }

                            cache.Append($"</{prefix}dataBar>");
                        }
                        else if (format.Type == eExcelConditionalFormattingRuleType.ThreeIconSet ||
                                 format.Type == eExcelConditionalFormattingRuleType.FourIconSet ||
                                 format.Type == eExcelConditionalFormattingRuleType.FiveIconSet)
                        {
                            var iconList = new List<ExcelConditionalFormattingIconDataBarValue>();
                            string iconSetString = "";
                            bool isCustom = false;
                            bool showValue;
                            bool reverse;
                            bool iconSetPercent;

                            switch (format.Type)
                            {
                                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                                    var threeIcon = (ExcelConditionalFormattingThreeIconSet)format;
                                    iconList.Add(threeIcon.Icon1);
                                    iconList.Add(threeIcon.Icon2);
                                    iconList.Add(threeIcon.Icon3);

                                    showValue = threeIcon.ShowValue;
                                    reverse = threeIcon.Reverse;
                                    iconSetPercent = threeIcon.IconSetPercent;
                                    isCustom = threeIcon.Custom;

                                    uid = threeIcon.Uid;
                                    iconSetString = threeIcon.GetIconSetString();
                                    break;

                                case eExcelConditionalFormattingRuleType.FourIconSet:
                                    var fourIcon = (ExcelConditionalFormattingFourIconSet)format;

                                    iconList.Add(fourIcon.Icon1);
                                    iconList.Add(fourIcon.Icon2);
                                    iconList.Add(fourIcon.Icon3);
                                    iconList.Add(fourIcon.Icon4);

                                    showValue = fourIcon.ShowValue;
                                    reverse = fourIcon.Reverse;
                                    iconSetPercent = fourIcon.IconSetPercent;
                                    isCustom = fourIcon.Custom;


                                    uid = fourIcon.Uid;
                                    iconSetString = fourIcon.GetIconSetString();
                                    break;

                                case eExcelConditionalFormattingRuleType.FiveIconSet:
                                    var fiveIcon = (ExcelConditionalFormattingFiveIconSet)format;

                                    iconList.Add(fiveIcon.Icon1);
                                    iconList.Add(fiveIcon.Icon2);
                                    iconList.Add(fiveIcon.Icon3);
                                    iconList.Add(fiveIcon.Icon4);
                                    iconList.Add(fiveIcon.Icon5);

                                    showValue = fiveIcon.ShowValue;
                                    reverse = fiveIcon.Reverse;
                                    iconSetPercent = fiveIcon.IconSetPercent;
                                    isCustom = fiveIcon.Custom;

                                    uid = fiveIcon.Uid;
                                    iconSetString = fiveIcon.GetIconSetString();
                                    break;
                                default:
                                    throw new InvalidOperationException($"Impossible case found {format.Type} is not an iconSet");
                            }

                            cache.Append($"<{prefix}cfRule type=\"iconSet\" priority=\"{format.Priority}\" id=\"{uid}\">");
                            cache.Append($"<{prefix}iconSet ");

                            if (isCustom == false)
                            {
                                cache.Append($"iconSet=\"{iconSetString}\"");
                            }

                            if (showValue == false)
                            {
                                cache.Append($" showValue=\"0\"");
                            }

                            if (iconSetPercent == false)
                            {
                                cache.Append($" percent=\"0\"");
                            }

                            if (reverse)
                            {
                                cache.Append($" reverse=\"1\"");
                            }

                            if (isCustom)
                            {
                                cache.Append(" custom=\"1\"");
                            }

                            cache.Append(">");

                            foreach (var icon in iconList)
                            {
                                cache.Append($"<{prefix}cfvo type=\"{icon.Type.ToString().UnCapitalizeFirstLetter()}\"");

                                if (icon.GreaterThanOrEqualTo == false && icon != iconList[0])
                                {
                                    cache.Append(" gte=\"0\"");
                                }

                                cache.Append(">");
                                if (icon.Formula != null)
                                {
                                    cache.Append($"<xm:f>{icon.Formula.EncodeXMLAttribute()}</xm:f>");
                                }
                                else if (icon.Value != double.NaN)
                                {
                                    cache.Append($"<xm:f>{icon.Value}</xm:f>");
                                }
                                cache.Append($"</{prefix}cfvo>");
                            }

                            if (isCustom)
                            {
                                for (int j = 0; j < iconList.Count; j++)
                                {
                                    string iconType = iconList[j].CustomIcon == null ? iconSetString : iconList[j].GetCustomIconStringValue();
                                    int iconIndex = iconList[j].CustomIcon == null ? j : iconList[j].GetCustomIconIndex();
                                    cache.Append($"<{prefix}cfIcon iconSet=\"{iconType}\" iconId=\"{iconIndex}\"/>");
                                }
                            }

                            cache.Append($"</{prefix}iconSet>");
                        }
                        else
                        {
                            cache.Append($"<{prefix}cfRule type=\"{format.GetAttributeType()}\"" +
                                $" priority=\"{format.Priority}\" ");

                            if (format.Operator != null)
                            {
                                cache.Append($"operator=\"{format.Operator.ToEnumString()}\" ");
                            }

                            cache.Append($"id=\"{format.Uid}\">");

                            if (format.Type == eExcelConditionalFormattingRuleType.TwoColorScale ||
                                format.Type == eExcelConditionalFormattingRuleType.ThreeColorScale)
                            {

                                cache.Append($"<{prefix}colorScale>");

                                var values = new List<ExcelConditionalFormattingColorScaleValue>()
                                    {format.As.TwoColorScale.LowValue, format.As.TwoColorScale.HighValue };

                                if (format.Type == eExcelConditionalFormattingRuleType.ThreeColorScale)
                                {
                                    values.Insert(1, format.As.ThreeColorScale.MiddleValue);
                                }

                                for (int j = 0; j < values.Count; j++)
                                {
                                    cache.Append($"<{prefix}cfvo type=\"{values[j].Type.ToEnumString()}\"");

                                    switch (values[j].Type)
                                    {
                                        case eExcelConditionalFormattingValueObjectType.Min:
                                        case eExcelConditionalFormattingValueObjectType.Max:
                                            cache.Append("/>");
                                            break;
                                        case eExcelConditionalFormattingValueObjectType.Formula:
                                            cache.Append(">");
                                            cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(values[j].Formula)}</xm:f>");
                                            cache.Append($"</{prefix}cfvo>");
                                            break;
                                        case eExcelConditionalFormattingValueObjectType.Percent:
                                        case eExcelConditionalFormattingValueObjectType.Percentile:
                                        case eExcelConditionalFormattingValueObjectType.Num:
                                            cache.Append(">");
                                            if (!string.IsNullOrEmpty(values[j].Formula))
                                            {
                                                cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(values[j].Formula)}</xm:f>");
                                            }
                                            else
                                            {
                                                cache.Append($"<xm:f>{values[j].Value}</xm:f>");
                                            }
                                            cache.Append($"</{prefix}cfvo>");
                                            break;
                                        default:
                                            throw new InvalidDataException(
                                                "WorksheetWriterXML detected unhandled eExcelConditionalFormattingValueObjectType");
                                    }
                                }

                                for (int j = 0; j < values.Count; j++)
                                {
                                    cache.Append($"<{prefix}color");

                                    if (values[j].ColorSettings.Theme != null)
                                    {
                                        cache.Append($" theme=\"{(int)values[j].ColorSettings.Theme}\"");
                                    }

                                    if (values[j].ColorSettings.Color != null && values[j].ColorSettings.Color != Color.Empty)
                                    {
                                        cache.Append($" rgb=\"FF{values[j].Color.ToColorString()}\"");
                                    }

                                    if (values[j].ColorSettings.Auto != null && values[j].ColorSettings.Auto != false)
                                    {
                                        cache.Append($" auto=\"1\"");
                                    }

                                    if (values[j].ColorSettings.Index != null)
                                    {
                                        cache.Append($" indexed=\"{values[j].ColorSettings.Index}\"");
                                    }

                                    if (values[j].ColorSettings.Tint != null)
                                    {
                                        cache.Append($" tint=\"{values[j].ColorSettings.Tint.Value.ToString(CultureInfo.InvariantCulture)}\"");
                                    }

                                    cache.Append($"/>");
                                }

                                cache.Append($"</{prefix}colorScale>");
                            }

                            if (!string.IsNullOrEmpty(format._formula))
                            {
                                cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(format._formula)}</xm:f>");
                            }

                            if (!string.IsNullOrEmpty(format._formula2))
                            {
                                cache.Append($"<xm:f>{ConvertUtil.ExcelEscapeAndEncodeString(format._formula2)}</xm:f>");
                            }

                            if (format.Style.HasValue)
                            {
                                WriteStyle(cache, prefix, format);
                            }
                        }
                        cache.Append($"</{prefix}cfRule>");

                        if (i == ruleList.Value.Count - 1)
                        {
                            if(format.Address == null)
                            {
                                cache.Append($"<xm:sqref></xm:sqref>");

                            }
                            else
                            {
                                cache.Append($"<xm:sqref>{format.Address.AddressSpaceSeparated}</xm:sqref>");
                            }
                            cache.Append($"</{prefix}conditionalFormatting>");
                        }
                    }
                }

                cache.Append($"</{prefix}conditionalFormattings>");
                cache.Append("</ext>");
            }
            return cache.ToString();
        }

        //TODO: Technically multiple attributes can exist at the same time. This doesn't support that.
        //But its also hard to imagine a scenario where both rgb and theme attributes could coexist.
        //Tint is supported to co-exist.
        string WriteColorOption(string nodeName, ExcelDxfColor color)
        {
            string returnString = "";

            if (color.Auto != null)
            {
                returnString = $"<{nodeName} auto=\"{(color.Auto == true ? 1 : 0)}\"";
            }

            if (color.Theme != null)
            {
                returnString = $"<{nodeName} theme=\"{(int)color.Theme}\"";
            }

            if(color.Color != null)
            { 
                Color baseColor = (Color)color.Color;
                returnString = $"<{nodeName} rgb=\"FF{baseColor.ToColorString()}\"";
            }

            if (color.Index != null)
            {
                returnString = $"<{nodeName} indexed=\"{color.Index}\"";
            }

            if (color.Tint != null)
            {
                returnString += string.Format(CultureInfo.InvariantCulture, " tint=\"{0}\"", color.Tint);
            }

            if(returnString != "")
            {
                returnString += "/>";
            }

            return returnString;
        }

        private string WriteCfIcon(ExcelConditionalFormattingIconDataBarValue icon, bool gteCheck = true)
        {
            StringBuilder cache = new StringBuilder();

            cache.Append($"<cfvo type=\"{icon.Type.ToString().UnCapitalizeFirstLetter()}\" ");

            if (icon.Formula != null)
            {
                cache.Append($"val=\"{icon.Formula.EncodeXMLAttribute()}\" ");
            }
            else if (icon.Value != double.NaN)
            {
                cache.Append($"val=\"{icon.Value}\" ");
            }

            if (icon.GreaterThanOrEqualTo == false && gteCheck == true)
            {
                cache.Append("gte=\"0\"");
            }

            cache.Append("/>");

            return cache.ToString();
        }

        private string UpdateConditionalFormattings(string prefix)
        {
            var cache = new StringBuilder();

            //Multiple cfRules may be applying to one address. Find out which
            var addressDict = new Dictionary<string, List<ExcelConditionalFormattingRule>>();

            foreach (var format in _ws.ConditionalFormatting)
            {
                if (format is ExcelConditionalFormattingRule cf)
                {
                    if (!cf.IsExtLst || cf.Type == eExcelConditionalFormattingRuleType.DataBar)
                    {
                        var key = cf.Address?.Address ?? "";
						if (addressDict.ContainsKey(key))
                        {
                            addressDict[key].Add(cf);
                        }
                        else
                        {
                            addressDict.Add(key, new List<ExcelConditionalFormattingRule>() { cf });
                        }
                    }
                }
            }

            //Iterate cfRules for each address
            foreach( var formatList in addressDict)
            {
                for (int i = 0; i < formatList.Value.Count; i++)
                {
                    var conditionalFormat = formatList.Value[i];

                    if(i == 0)
                    {
						cache.Append($"<conditionalFormatting");
						if (conditionalFormat.Address!=null)
                        {
							cache.Append($" sqref=\"{conditionalFormat.Address.AddressSpaceSeparated}\"");
						}
						if (conditionalFormat.PivotTable)
                        {
                            cache.Append($" pivot=\"1\"");
                        }
                        cache.Append($">");
                    }

                    cache.Append($"<cfRule type=\"{conditionalFormat.GetAttributeType()}\" ");

                    cache.Append(UpdateConditionalFormattingAttributes(conditionalFormat));
                    cache.Append($">");

                    if (string.IsNullOrEmpty(conditionalFormat._formula) == false)
                    {
                        cache.Append("<formula>" + ConvertUtil.ExcelEscapeAndEncodeString(conditionalFormat._formula) + "</formula>");
                        if (string.IsNullOrEmpty(conditionalFormat._formula2) == false)
                        {
                            cache.Append("<formula>" + ConvertUtil.ExcelEscapeAndEncodeString(conditionalFormat._formula2) + "</formula>");
                        }
                    }

                    if (conditionalFormat.Type == eExcelConditionalFormattingRuleType.TwoColorScale ||
                        conditionalFormat.Type == eExcelConditionalFormattingRuleType.ThreeColorScale)
                    {
                        cache.Append("<colorScale>");

                        var colorScaleValues = new List<ExcelConditionalFormattingColorScaleValue>()
                        {
                            ((ExcelConditionalFormattingTwoColorScale)conditionalFormat).LowValue,
                            ((ExcelConditionalFormattingTwoColorScale)conditionalFormat).HighValue
                        };

                        if (conditionalFormat.Type == eExcelConditionalFormattingRuleType.ThreeColorScale)
                        {
                            colorScaleValues.Insert(1, conditionalFormat.As.ThreeColorScale.MiddleValue);
                        }

                        for(int j = 0; j < colorScaleValues.Count; j++)
                        {
                            var cSValue = colorScaleValues[j];

                            cache.Append($"<cfvo type=\"{cSValue.Type.ToString().UnCapitalizeFirstLetter()}\" ");

                            if (!double.IsNaN(cSValue.Value))
                            {
                                cache.Append($"val=\"{cSValue.Value}\"");
                            }
                            else if (!string.IsNullOrEmpty(cSValue.Formula))
                            {
                                cache.Append($"val=\"{cSValue.Formula.EncodeXMLAttribute()}\"");
                            }

                            //Note: No GTE bool attribute as it is only applicable to iconsets according to Microsofts documentation.

                            cache.Append("/>");
                        }

                        for (int j = 0; j < colorScaleValues.Count; j++)
                        {
                            var cSValue = colorScaleValues[j];

                            cache.Append($"<color");

                            if (cSValue.ColorSettings.Auto != null && cSValue.ColorSettings.Auto != false)
                            {
                                cache.Append($" auto=\"1\"");
                            }

                            if (cSValue.ColorSettings.Index != null)
                            {
                                cache.Append($" indexed=\"{cSValue.ColorSettings.Index}\"");
                            }

                            if (cSValue.ColorSettings.Color != null && cSValue.ColorSettings.Color != Color.Empty)
                            {
                                cache.Append($" rgb=\"FF{cSValue.Color.ToColorString()}\"");
                            }

                            if (cSValue.ColorSettings.Theme != null)
                            {
                                cache.Append($" theme=\"{(int)cSValue.ColorSettings.Theme}\"");
                            }

                            if (cSValue.ColorSettings.Tint != null)
                            {
                                cache.Append($" tint=\"{cSValue.ColorSettings.Tint.Value.ToString(CultureInfo.InvariantCulture)}\"");
                            }

                            cache.Append($"/>");
                        }

                        cache.Append("</colorScale>");
                    }

                    if (conditionalFormat.IsExtLst)
                    {
                        if (conditionalFormat.Type == eExcelConditionalFormattingRuleType.DataBar)
                        {
                            var dataBar = (ExcelConditionalFormattingDataBar)conditionalFormat;
                            cache.Append($"<dataBar>");

                            string typeLow = dataBar.LowValue.Type.ToString().UnCapitalizeFirstLetter();
                            string typeHigh = dataBar.HighValue.Type.ToString().UnCapitalizeFirstLetter();

                            if (typeLow.Contains("auto"))
                            {
                                typeLow = typeLow.Remove(0,"auto".Length);
                            }

                            if (typeHigh.Contains("auto"))
                            {
                                typeHigh = typeHigh.Remove(0, "auto".Length);
                            }

                            cache.Append($"<cfvo type=\"{typeLow.UnCapitalizeFirstLetter()}\"");

                            if (dataBar.LowValue.HasValueOrFormula)
                            {
                                if (!string.IsNullOrEmpty(dataBar.LowValue.Formula))
                                {
                                    cache.Append($" val=\"{dataBar.LowValue.Formula.EncodeXMLAttribute()}\"");
                                }
                                else
                                {
                                    cache.Append($" val=\"{dataBar.LowValue.Value}\"");
                                }
                            }

                            cache.Append("/>");

                            cache.Append($"<cfvo type=\"{typeHigh.UnCapitalizeFirstLetter()}\"");

                            if (dataBar.HighValue.HasValueOrFormula)
                            {
                                if (!string.IsNullOrEmpty(dataBar.HighValue.Formula))
                                {
                                    cache.Append($" val=\"{dataBar.HighValue.Formula.EncodeXMLAttribute()}\"");
                                }
                                else
                                {
                                    cache.Append($" val=\"{dataBar.HighValue.Value}\"");
                                }
                            }

                            cache.Append("/>");

                            WriteDxfColor("", cache, dataBar.FillColor);

                            cache.Append($"</dataBar>");

                            cache.Append($"<extLst>");

                            prefix = "x14";
                            cache.Append($"<ext xmlns:{prefix}=\"{ExcelPackage.schemaMainX14}\" uri=\"{ExtLstUris.ExtChildUri}\">");
                            cache.Append($"<{prefix}:id>{dataBar.Uid}</{prefix}:id>");
                            cache.Append($"</ext>");

                            cache.Append($"</extLst>");
                        }
                    }
                    else if (conditionalFormat.Type == eExcelConditionalFormattingRuleType.ThreeIconSet ||
                             conditionalFormat.Type == eExcelConditionalFormattingRuleType.FourIconSet ||
                             conditionalFormat.Type == eExcelConditionalFormattingRuleType.FiveIconSet)
                    {

                        var iconList = new List<ExcelConditionalFormattingIconDataBarValue>();
                        string iconSetString = "";
                        bool showValue = false;
                        bool reverse = false;
                        bool iconSetPercent = true;

                        switch (conditionalFormat.Type)
                        {
                            case eExcelConditionalFormattingRuleType.ThreeIconSet:
                                var threeIcon = (ExcelConditionalFormattingThreeIconSet)conditionalFormat;
                                iconList.Add(threeIcon.Icon1);
                                iconList.Add(threeIcon.Icon2);
                                iconList.Add(threeIcon.Icon3);

                                showValue = threeIcon.ShowValue;
                                reverse = threeIcon.Reverse;
                                iconSetPercent = threeIcon.IconSetPercent;

                                iconSetString = threeIcon.GetIconSetString();
                                break;

                            case eExcelConditionalFormattingRuleType.FourIconSet:
                                var fourIcon = (ExcelConditionalFormattingFourIconSet)conditionalFormat;
                                iconList.Add(fourIcon.Icon1);
                                iconList.Add(fourIcon.Icon2);
                                iconList.Add(fourIcon.Icon3);
                                iconList.Add(fourIcon.Icon4);

                                showValue = fourIcon.ShowValue;
                                reverse = fourIcon.Reverse;
                                iconSetPercent = fourIcon.IconSetPercent;

                                iconSetString = fourIcon.GetIconSetString();
                                break;

                            case eExcelConditionalFormattingRuleType.FiveIconSet:
                                var fiveIcon = (ExcelConditionalFormattingFiveIconSet)conditionalFormat;

                                iconList.Add(fiveIcon.Icon1);
                                iconList.Add(fiveIcon.Icon2);
                                iconList.Add(fiveIcon.Icon3);
                                iconList.Add(fiveIcon.Icon4);
                                iconList.Add(fiveIcon.Icon5);

                                showValue = fiveIcon.ShowValue;
                                reverse = fiveIcon.Reverse;
                                iconSetPercent = fiveIcon.IconSetPercent;

                                iconSetString = fiveIcon.GetIconSetString();
                                break;
                        }

                        cache.Append($"<iconSet iconSet=\"{iconSetString}\"");

                        if (showValue == false)
                        {
                            cache.Append($" showValue=\"0\"");
                        }

                        if(iconSetPercent == false)
                        {
                            cache.Append($" percent=\"0\"");
                        }

                        if (reverse)
                        {
                            cache.Append($" reverse=\"1\"");
                        }

                        cache.Append(">");

                        for (int j = 0; j < iconList.Count; j++)
                        {
                            if (j == 0)
                            {
                                cache.Append(WriteCfIcon(iconList[j], false));

                            }
                            else
                            {
                                cache.Append(WriteCfIcon(iconList[j]));
                            }
                        }

                        cache.Append($"</iconSet>");
                    }
                    cache.Append($"</cfRule>");
                }
                cache.Append($"</conditionalFormatting>");
            }

            return cache.ToString();
        }

        void WriteDxfColor(string prefix, StringBuilder cache, ExcelDxfColor col, string nodeName = "color")
        {
            cache.Append($"<{prefix}{nodeName}");

            if (col.Theme != null)
            {
                cache.Append($" theme=\"{(int)col.Theme}\"");
            }

            if (col.Color != null && col.Color != Color.Empty)
            {
                cache.Append($" rgb=\"FF{((Color)col.Color).ToColorString()}\"");
            }

            if (col.Auto != null && col.Auto != false)
            {
                cache.Append($" auto=\"1\"");
            }

            if (col.Index != null)
            {
                cache.Append($" indexed=\"{col.Index}\"");
            }

            if (col.Tint != null)
            {
                cache.Append($" tint=\"{col.Tint.Value.ToString(CultureInfo.InvariantCulture)}\"");
            }

            cache.Append($"/>");
        }

        void WriteStyle(StringBuilder cache, string prefix, ExcelConditionalFormattingRule format)
        {
            if (format.Style.HasValue)
            {
                cache.Append($"<{prefix}dxf>");

                if (format.Style.Font.HasValue)
                {
                    cache.Append($"<font>");

                    if (format.Style.Font.Bold == true)
                    {
                        cache.Append($"<b/>");
                    }

                    if (format.Style.Font.Italic == true)
                    {
                        if (format.Style.Font.Bold == false || format.Style.Font.Bold == null)
                        {
                            cache.Append("<b val=\"0\"/>");
                        }
                        cache.Append($"<i/>");
                    }

                    if (format.Style.Font.Bold == false && format.Style.Font.Italic == false)
                    {
                        cache.Append("<b val=\"0\"/>");
                        cache.Append("<i val=\"0\"/>");
                    }

                    if (format.Style.Font.Strike == true)
                    {
                        cache.Append($"<strike/>");

                    }

                    if (format.Style.Font.Underline.HasValue == true)
                    {
                        cache.Append($"<u");
                        if (format.Style.Font.Underline.Value == Style.ExcelUnderLineType.Double)
                        {
                            cache.Append(" val=\"double\"");
                        }
                        cache.Append($"/>");
                    }

                    if (format.Style.Font.Color.HasValue == true)
                    {
                        cache.Append(WriteColorOption("color", format.Style.Font.Color));
                    }

                    cache.Append("</font>");
                }

                if (format.Style.NumberFormat.HasValue)
                {
                    var numberFormat = format.Style.NumberFormat.Format.EncodeXMLAttribute();

                    cache.Append($"<numFmt numFmtId =\"{format.Style.NumberFormat.NumFmtID}\" " +
                        $"formatCode = \"{numberFormat}\"/>");
                }

                if (format.Style.Fill.HasValue)
                {
                    cache.Append("<fill>");

                    switch (format.Style.Fill.Style)
                    {
                        case Style.eDxfFillStyle.PatternFill:

                            if (format.Style.Fill.PatternType != null && format.Style.Fill.PatternType != ExcelFillStyle.None)
                            {
                                cache.Append($"<patternFill patternType=\"{format.Style.Fill.PatternType.ToEnumString()}\">");
                            }
                            else
                            {
                                cache.Append("<patternFill>");
                            }

                            if (format.Style.Fill.PatternColor.HasValue)
                            {
                                cache.Append(WriteColorOption("fgColor", format.Style.Fill.PatternColor));
                            }

                            if (format.Style.Fill.BackgroundColor.HasValue)
                            {
                                cache.Append(WriteColorOption("bgColor", format.Style.Fill.BackgroundColor));
                            }

                            cache.Append("</patternFill>");
                            break;
                        case Style.eDxfFillStyle.GradientFill:

                            cache.Append("<gradientFill");

                            if (format.Style.Fill.Gradient.GradientType != null)
                            {
                                cache.Append($" type=\"{format.Style.Fill.Gradient.GradientType.ToString().ToLower()}\"");
                            }

                            if (format.Style.Fill.Gradient.Degree != null)
                            {
                                cache.Append(string.Format(CultureInfo.InvariantCulture, " degree=\"{0}\"", format.Style.Fill.Gradient.Degree.Value));
                            }

                            if (format.Style.Fill.Gradient.Left != null)
                            {
                                cache.Append(string.Format(CultureInfo.InvariantCulture, " left=\"{0}\"", format.Style.Fill.Gradient.Left.Value));
                            }

                            if (format.Style.Fill.Gradient.Right != null)
                            {
                                cache.Append(string.Format(CultureInfo.InvariantCulture, " right=\"{0}\"", format.Style.Fill.Gradient.Right.Value));
                            }

                            if (format.Style.Fill.Gradient.Top != null)
                            {
                                cache.Append(string.Format(CultureInfo.InvariantCulture, " top=\"{0}\"", format.Style.Fill.Gradient.Top.Value));
                            }

                            if (format.Style.Fill.Gradient.Bottom != null)
                            {
                                cache.Append(string.Format(CultureInfo.InvariantCulture, " bottom=\"{0}\"", format.Style.Fill.Gradient.Bottom.Value));
                            }

                            cache.Append(">");
                            cache.Append("<stop position=\"0\">");

                            cache.Append(WriteColorOption("color", format.Style.Fill.Gradient.Colors[0].Color));

                            cache.Append("</stop>");
                            cache.Append("<stop position=\"1.0\">");

                            cache.Append(WriteColorOption("color", format.Style.Fill.Gradient.Colors[1].Color));

                            cache.Append("</stop>");

                            cache.Append("</gradientFill>");

                            break;
                    }
                    cache.Append("</fill>");
                }

                if (format.Style.Border.HasValue)
                {
                    cache.Append("<border>");

                    if (format.Style.Border.Left.HasValue)
                    {
                        cache.Append($"<left style=\"{format.Style.Border.Left.Style.ToString().ToLower()}\">");
                        cache.Append(WriteColorOption("color", format.Style.Border.Left.Color));
                        cache.Append("</left>");
                    }

                    if (format.Style.Border.Right.HasValue)
                    {
                        cache.Append($"<right style=\"{format.Style.Border.Right.Style.ToString().ToLower()}\">");
                        cache.Append(WriteColorOption("color", format.Style.Border.Right.Color));
                        cache.Append("</right>");
                    }

                    if (format.Style.Border.Top.HasValue)
                    {
                        cache.Append($"<top style=\"{format.Style.Border.Top.Style.ToString().ToLower()}\">");
                        cache.Append(WriteColorOption("color", format.Style.Border.Top.Color));
                        cache.Append("</top>");
                    }

                    if (format.Style.Border.Bottom.HasValue)
                    {
                        cache.Append($"<bottom style=\"{format.Style.Border.Bottom.Style.ToString().ToLower()}\">");
                        cache.Append(WriteColorOption("color", format.Style.Border.Bottom.Color));
                        cache.Append("</bottom>");
                    }

                    cache.Append("<vertical/>");
                    cache.Append("<horizontal/>");
                    cache.Append("</border>");
                }

                cache.Append($"</{prefix}dxf>");
            }
        }
    }
}
