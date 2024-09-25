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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core
{
    internal class RangeCopyHelper
    {
        private class CopiedCell
        {
            internal int Row { get; set; }
            internal int Column { get; set; }
            internal object Value { get; set; }
            internal string Type { get; set; }
            internal object Formula { get; set; }
            internal int? StyleID { get; set; }
            internal Uri HyperLink { get; set; }
            internal ExcelComment Comment { get; set; }
            internal ExcelThreadedCommentThread ThreadedComment { get; set; }
            internal byte Flag { get; set; }
            internal ExcelWorksheet.MetaDataReference MetaData{ get; set; }
    }
        private readonly ExcelRangeBase _sourceRange;
        private readonly ExcelRangeBase _destinationRange;
        private readonly ExcelRangeCopyOptionFlags _copyOptions;
        private readonly bool _sameWorkbook;
		private ExcelMetadata _sourceMd, _destMd;
		Dictionary<ulong, CopiedCell> _copiedCells=new Dictionary<ulong, CopiedCell>();
        int _sourceDaIx = -1;
        int _destDaIx = -1;
        internal RangeCopyHelper(ExcelRangeBase sourceRange, ExcelRangeBase destination, ExcelRangeCopyOptionFlags copyOptions)
        {
            _sourceRange = sourceRange;
            _destinationRange = destination;
			_sameWorkbook = _destinationRange._worksheet.Workbook == _sourceRange._worksheet.Workbook;
			_copyOptions = copyOptions;
        }
        internal void Copy()
        {
            GetCopiedValues();

            Dictionary<int, ExcelAddress> copiedMergedCells;
            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeMergedCells))
            {
                copiedMergedCells = GetCopiedMergedCells();
            }
            else
            {
                copiedMergedCells = null;
            }
            
            ClearDestination();

            CopyValuesToDestination();
            
            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeDataValidations))
            {
                CopyDataValidations();
            }

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting))
            {
                CopyConditionalFormatting();
            }

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeMergedCells))
            {
                CopyMergedCells(copiedMergedCells);
            }

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeDrawings))
            {
                CopyDrawings();
            }

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeTables))
            {
                CopyTables();
            }

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludePivotTables))
            {
                CopyPivotTables();
            }

            CopyFullColumn();
            CopyFullRow();
        }

        private void CopyPivotTables()
        {
            var tablesToCopy = new List<ExcelPivotTable>();
            foreach (var table in _sourceRange.Worksheet.PivotTables)
            {
                var ac = _sourceRange.Collide(table.Address);
                if (ac == eAddressCollition.Inside ||
                    ac == eAddressCollition.Equal)
                {
                    tablesToCopy.Add(table);
                }
            }
            tablesToCopy.ForEach(table => CopyPivotTable(table));
        }

        private void CopyTables()
        {
            var tablesToCopy = new List<ExcelTable>();
            foreach(var table in _sourceRange.Worksheet.Tables)
            {
                var ac = _sourceRange.Collide(table.Range);
                if (ac == eAddressCollition.Inside ||
                    ac == eAddressCollition.Equal)
                {
                    tablesToCopy.Add(table); 
                }
            }
            tablesToCopy.ForEach(table=> CopyTable(table));
        }

        private void CopyTable(ExcelTable table)
        {
            var tr = table.Range;
            var dr = _destinationRange;
            var copiedTable = _destinationRange.Worksheet.Cells[
                dr._fromRow + (tr._fromRow - _sourceRange._fromRow),
                dr._fromCol + (tr._fromCol -_sourceRange._fromCol),
                dr._fromRow + (tr._toRow - _sourceRange._fromRow),
                dr._fromCol + (tr._toCol - _sourceRange._fromCol)];

            var name = table.Name;
            if (_destinationRange._workbook.ExistsTableName(name))
            {
                name = _destinationRange.Worksheet.Tables.GetNewTableName(name);
            }
            _destinationRange._worksheet.Tables.AddInternal(copiedTable, name, table);
        }
        private void CopyPivotTable(ExcelPivotTable ptCopy)
        {
            var tr = ptCopy.Address;
            var dr = _destinationRange;
            var destinationAddress = _destinationRange.Worksheet.Cells[
                dr._fromRow + (tr._fromRow - _sourceRange._fromRow),
                dr._fromCol + (tr._fromCol - _sourceRange._fromCol),
                dr._fromRow + (tr._toRow - _sourceRange._fromRow),
                dr._fromCol + (tr._toCol - _sourceRange._fromCol)];

            var name = ptCopy.Name;
            if (_destinationRange.Worksheet.PivotTables._pivotTableNames.ContainsKey(name))
            {
                name = _destinationRange.Worksheet.PivotTables.GetNewTableName(name);
            }

            _destinationRange._worksheet.PivotTables.Add(new ExcelPivotTable(_destinationRange.Worksheet, destinationAddress, ptCopy, name, _destinationRange.Worksheet.Workbook._nextPivotTableID++));
        }

        private void CopyDrawings()
        {
            foreach(var drawing in _sourceRange._worksheet.Drawings.ToList())
            {
                var drawingRange = new ExcelAddress(drawing.From.Row+1, drawing.From.Column+1, drawing.To.Row+1, drawing.To.Column + 1);
                if (_sourceRange.Intersect(drawingRange) != null )
                {
                    var row = drawingRange._fromRow - _sourceRange._fromRow;
                    row = _destinationRange._fromRow + row - 1;
                    var col = drawingRange._fromCol - _sourceRange._fromCol;
                    col = _destinationRange._fromCol + col - 1;
                    drawing.Copy(_destinationRange.Worksheet, row, col);
                }
            }
        }

        private void CopyDataValidations()
        {
            foreach (var idv in _sourceRange._worksheet.DataValidations)
            {
                if (idv is ExcelDataValidation dv)
                {
                    string newAddress = "";
                    if (dv.Address.Addresses == null)
                    {
                        newAddress = HandelAddress(dv.Address);
                    }
                    else
                    {
                        foreach (var a in dv.Address.Addresses)
                        {
                            var na = HandelAddress(a);
                            if (!string.IsNullOrEmpty(na))
                            {
                                if (string.IsNullOrEmpty(newAddress))
                                {
                                    newAddress += na;
                                }
                                else
                                {
                                    newAddress += "," + na;
                                }

                            }
                        }
                    }

                    if (string.IsNullOrEmpty(newAddress) == false)
                    {
                        if (_sourceRange._worksheet == _destinationRange._worksheet)
                        {
                            dv.SetAddress(dv.Address + "," + newAddress);
                            dv._ws.DataValidations.UpdateRangeDictionary(dv);
                        }
                        else
                        {
                            _destinationRange._worksheet.DataValidations.AddCopyOfDataValidation(dv, _destinationRange._worksheet, new ExcelAddressBase(newAddress).AddressSpaceSeparated);
                        }
                    }
                }
            }
        }

        private void CopyConditionalFormatting()
        {
            foreach(var cf in _sourceRange._worksheet.ConditionalFormatting)
            {
                string newAddress = "";
                if (cf.Address.Addresses==null)
                {
                    newAddress = HandelAddress(cf.Address);
                }
                else
                {
                    foreach (var a in cf.Address.Addresses)
                    {
                        var na = HandelAddress(a);
                        if(!string.IsNullOrEmpty(na))
                        {
                            if(string.IsNullOrEmpty(newAddress))
                            {
                                newAddress += na;
                            }
                            else
                            {
                                newAddress += "," + na ;
                            }
                            
                        }
                    }
                }

                if (string.IsNullOrEmpty(newAddress) == false)
                {
                    if (_sourceRange._worksheet == _destinationRange._worksheet)
                    {
                        cf.Address = new ExcelAddress(cf.Address + "," + newAddress);
                    }
                    else
                    {
                        _destinationRange._worksheet.ConditionalFormatting.CopyRule((ExcelConditionalFormattingRule)cf, new ExcelAddress(newAddress));
                        if (cf.Style.HasValue)
                        {
                            var destRule = ((ExcelConditionalFormattingRule)_destinationRange._worksheet.ConditionalFormatting[_destinationRange._worksheet.ConditionalFormatting.Count - 1]);
                            destRule.SetStyle((ExcelDxfStyleConditionalFormatting)cf.Style.Clone());
                        }
                    }
                }
            }
        }

        private string HandelAddress(ExcelAddressBase cfAddress)
        {
            if (cfAddress.Collide(_sourceRange) != eAddressCollition.No)
            {
                var address = _sourceRange.Intersect(cfAddress);
                var rowOffset = address._fromRow - _sourceRange._fromRow;
                var colOffset = address._fromCol - _sourceRange._fromCol;
                var fr = Math.Min(Math.Max(_destinationRange._fromRow + rowOffset, 1), ExcelPackage.MaxRows);
                var fc = Math.Min(Math.Max(_destinationRange._fromCol + colOffset, 1), ExcelPackage.MaxColumns);
                address = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? new ExcelAddressBase(fc, fr, Math.Min(fc + address.Columns - 1, ExcelPackage.MaxColumns), Math.Min(fr + address.Rows - 1, ExcelPackage.MaxRows)) :
                                                                                                new ExcelAddressBase(fr, fc, Math.Min(fr + address.Rows-1, ExcelPackage.MaxRows), Math.Min(fc + address.Columns-1, ExcelPackage.MaxColumns));
                return address.Address;
            }
            return "";
        }

        private void GetCopiedValues()
        {
            var worksheet = _sourceRange._worksheet;
            var toRow = _sourceRange._toRow;
            var toCol = _sourceRange._toCol;
            var fromRow = _sourceRange._fromRow;
            var fromCol = _sourceRange._fromCol;

            var includeValues = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues);
            var includeStyles = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles);
            var includeComments = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments);
            var includeThreadedComments = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments);

            Dictionary<int, int> styleCashe = new Dictionary<int, int>();

            AddValuesFormulasAndStyles(worksheet, includeStyles, styleCashe);

            if (includeComments)
            {
                AddComments(worksheet);
            }

            if (includeThreadedComments)
            {
                AddThreadedComments(worksheet);
            }
        }

        private void AddValuesFormulasAndStyles(ExcelWorksheet worksheet, bool includeStyles, Dictionary<int, int> styleCashe)
        {
            int styleId = 0;
            object o = null;
            byte flag = 0;
            Uri hl = null;
            var includeValues = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues);
            var includeFormulas = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas);
            var includeHyperlinks = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks);
            if (includeValues == false && includeHyperlinks == false && includeFormulas == false) return;
            var cse = new CellStoreEnumerator<ExcelValue>(worksheet._values,  _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cell = new CopiedCell
                {
                    Row = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromRow + (col - _sourceRange._fromCol) : _destinationRange._fromRow + (row - _sourceRange._fromRow),
                    Column = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromCol + (row - _sourceRange._fromRow) : _destinationRange._fromCol + (col - _sourceRange._fromCol),
                };

                if(includeValues)
                {
                    cell.Value = cse.Value._value;
                }

                if (includeFormulas && worksheet._formulas.Exists(row, col, ref o))
                {
                    if (o is int)
                    {
                        cell.Formula = worksheet.GetFormula(cse.Row, cse.Column);
                        if (worksheet._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.ArrayFormula))
                        {
                            _destinationRange._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
                        }
                        // We currently don't copy CellFlags.DataTableFormula's, as Excel does not.
                    }
                    else
                    {
                        cell.Formula = o;
                    }
                }

                if (includeStyles && worksheet.ExistsStyleInner(row, col, ref styleId))
                {
                    if (_sameWorkbook)
                    {
                        cell.StyleID = styleId;
                    }
                    else
                    {
                        if (styleCashe.ContainsKey(styleId))
                        {
                            styleId = styleCashe[styleId];
                        }
                        else
                        {
                            var oldStyleID = styleId;
                            styleId = _destinationRange._worksheet.Workbook.Styles.CloneStyle(_sourceRange._worksheet.Workbook.Styles, styleId);
                            styleCashe.Add(oldStyleID, styleId);
                        }
                        cell.StyleID = styleId;
                    }
                }

                var md = new ExcelWorksheet.MetaDataReference();
                if (includeFormulas && worksheet._metadataStore.Exists(row, col, ref md))
                {
                    cell.MetaData = md;
                }

                if (includeHyperlinks && worksheet._hyperLinks.Exists(row, col, ref hl))
                {
                    cell.HyperLink = hl;
                }

                if (worksheet._flags.Exists(row, col, ref flag))
                {
                    cell.Flag = flag;
                }

                _copiedCells.Add(ExcelCellBase.GetCellId(0, row, col), cell);
            }
        }
        private void AddComments(ExcelWorksheet worksheet)
        {
            var cse = new CellStoreEnumerator<int>(worksheet._commentsStore, _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cellId = ExcelCellBase.GetCellId(0, row, col);
                CopiedCell cell;
                if (_copiedCells.ContainsKey(cellId))
                {
                    cell = _copiedCells[cellId];
                }
                else
                {
                    cell = new CopiedCell
                    {
                        Row = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromRow + (col - _sourceRange._fromCol) : _destinationRange._fromRow + (row - _sourceRange._fromRow),
                        Column = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromCol + (row - _sourceRange._fromRow) : _destinationRange._fromCol + (col - _sourceRange._fromCol),
                    };
                    _copiedCells.Add(cellId, cell);
                }
                cell.Comment = worksheet._comments[cse.Value];
            }
        }
        private void AddThreadedComments(ExcelWorksheet worksheet)
        {
            var cse = new CellStoreEnumerator<int>(worksheet._threadedCommentsStore, _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cellId = ExcelCellBase.GetCellId(0, row, col);
                CopiedCell cell;
                if (_copiedCells.ContainsKey(cellId))
                {
                    cell = _copiedCells[cellId];
                }
                else
                {
                    cell = new CopiedCell
                    {
                        Row = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromRow + (col - _sourceRange._fromCol) : _destinationRange._fromRow + (row - _sourceRange._fromRow),
                        Column = EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose) ? _destinationRange._fromCol + (row - _sourceRange._fromRow) : _destinationRange._fromCol + (col - _sourceRange._fromCol),
                    };
                    _copiedCells.Add(cellId, cell);
                }
                cell.ThreadedComment = worksheet._threadedComments[cse.Value];
            }
        }
        private void CopyValuesToDestination()
        {
            int fromRow = _sourceRange._fromRow;
            int fromCol = _sourceRange._fromCol;
            foreach (var cell in _copiedCells.Values)
            {
                if (EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues) && 
                    EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
                {
                    _destinationRange._worksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID ?? 0);
                }
                else if(EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
                {
                    _destinationRange._worksheet.SetValueInner(cell.Row, cell.Column, cell.Value);
                }
                else
                {
                    _destinationRange._worksheet.SetValueStyleIdInner(cell.Row, cell.Column, cell.Value, cell.StyleID ?? 0);
                }
                if(cell.Value is ExcelRichTextCollection)
                {
                    var t = new ExcelRichTextCollection((Style.ExcelRichTextCollection)cell.Value, _destinationRange);
                    _destinationRange._worksheet.SetValueInner(cell.Row, cell.Column,t);
                }

                if ((EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas) && EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues)) &&
                    cell.Formula != null)
                {
                    cell.Formula = ExcelRangeBase.UpdateFormulaReferences(cell.Formula.ToString(), _destinationRange._fromRow - fromRow, _destinationRange._fromCol - fromCol, 0, 0, _destinationRange.WorkSheetName, _destinationRange.WorkSheetName, true, true);
                    _destinationRange._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }

                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks) && 
                    cell.HyperLink != null)
                {
                    _destinationRange._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
                }

                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments) && 
                    cell.ThreadedComment != null)
                {
                    var differentPackages = _destinationRange._workbook != _sourceRange._workbook;
                    var tc = _destinationRange.Worksheet.Cells[cell.Row, cell.Column].AddThreadedComment();
                    foreach (var c in cell.ThreadedComment.Comments)
                    {
                        if(differentPackages && _destinationRange._workbook.ThreadedCommentPersons[c.PersonId]==null)
                        {
                            var p = _sourceRange._workbook.ThreadedCommentPersons[c.PersonId];
                            _destinationRange._workbook.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                        }
                        tc.AddCommentFromXml((XmlElement)c.TopNode);
                    }
                }
                else if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments) && 
                    cell.Comment != null)
                {
                    CopyComment(_destinationRange, cell);
                }

                if (cell.Flag != 0)
                {
                    _destinationRange._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
                }

                if(EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas | ExcelRangeCopyOptionFlags.ExcludeValues) &&
                    cell.MetaData.cm > 0 || cell.MetaData.vm > 0)
                {
					if (_sameWorkbook == false)
					{
						CopyMetaDataToNewPackage(cell);
					}
					_destinationRange._worksheet._metadataStore.SetValue(cell.Row, cell.Column, cell.MetaData);
                }
            }
        }

		private void CopyMetaDataToNewPackage(CopiedCell cell)
		{
            if(cell.MetaData.cm > 0)
            {
                if(_sourceDaIx==-1)
                {
                    _sourceMd = _sourceRange.Worksheet.Workbook.Metadata;
					_destMd = _destinationRange.Worksheet.Workbook.Metadata;
                    _sourceMd.GetDynamicArrayIndex(out _sourceDaIx);
				}

				var md = cell.MetaData;
				if (cell.MetaData.cm == _sourceDaIx)
                {
                    if(_destDaIx < 0)
                    {
						_destMd.GetDynamicArrayIndex(out _destDaIx);
					}
                    md.cm = _destDaIx;
                    cell.MetaData = md;
				}
                else
                {
					cell.MetaData = default;
				}
			}
            else
            {
                //We don't copy value meta data. Errors are handled on save for rich data types like #CALC and #SPILL, via the error values. Rich Data - DataTypes are currently not supported.
                cell.MetaData = default;
            }
		}

		private static void CopyComment(ExcelRangeBase destination, CopiedCell cell)
        {
            var c = destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment.Text, cell.Comment.Author);
            var offsetCol = c.Column - cell.Comment.Column;
            var offsetRow = c.Row - cell.Comment.Row;
            XmlHelper.CopyElement((XmlElement)cell.Comment.TopNode, (XmlElement)c.TopNode, new string[] { "id", "spid" });

            if (c.From.Column + offsetCol >= 0)
            {
                c.From.Column += offsetCol;
                c.To.Column += offsetCol;
            }
            if (c.From.Row + offsetRow >= 0)
            {
                c.From.Row += offsetRow;
                c.To.Row += offsetRow;
            }
            c.Row = cell.Row-1;
            c.Column = cell.Column-1;

            c._commentHelper.TopNode.InnerXml = cell.Comment._commentHelper.TopNode.InnerXml;
            c.RichText = new Style.ExcelRichTextCollection(cell.Comment.RichText, destination);
            //Add relation to image used for filling the comment
            if(cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Frame ||
              cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Tile ||
              cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Pattern)
            {
                var img = cell.Comment.Fill.PatternPictureSettings.Image;
                if (img.ImageBytes != null)
                {
                    c.Fill.PatternPictureSettings.Image.SetImage(img.ImageBytes, img.Type ?? ePictureType.Jpg);
                }
            }
        }

        private void ClearDestination()
        {
            //Clear all existing cells; 
            int rows = _sourceRange._toRow - _sourceRange._fromRow + 1,
                cols = _sourceRange._toCol - _sourceRange._fromCol + 1;
            if (EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose))
            {
                rows = _sourceRange._toCol - _sourceRange._fromCol + 1;
                cols = _sourceRange._toRow - _sourceRange._fromRow + 1;
            }

            _destinationRange._worksheet.MergedCells.Clear(new ExcelAddressBase(_destinationRange._fromRow, _destinationRange._fromCol, _destinationRange._fromRow + rows - 1, _destinationRange._fromCol + cols - 1));

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues) && EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
            {
                _destinationRange._worksheet._values.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            }
            _destinationRange._worksheet._formulas.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            _destinationRange._worksheet._metadataStore.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            _destinationRange._worksheet._hyperLinks.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            _destinationRange._worksheet._flags.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            _destinationRange._worksheet._commentsStore.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
            _destinationRange._worksheet._threadedCommentsStore.Clear(_destinationRange._fromRow, _destinationRange._fromCol, rows, cols);
        }

        private Dictionary<int, ExcelAddress> GetCopiedMergedCells()
        {
            var worksheet = _sourceRange._worksheet;
            var copiedMergedCells = new Dictionary<int, ExcelAddress>();
            //Merged cells
            var csem = new CellStoreEnumerator<int>(worksheet.MergedCells._cells, _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            while (csem.Next())
            {
                if (!copiedMergedCells.ContainsKey(csem.Value))
                {
                    var adr = new ExcelAddress(worksheet.Name, worksheet.MergedCells._list[csem.Value]);
                    var collideResult = _sourceRange.Collide(adr);
                    if (collideResult == eAddressCollition.Inside || collideResult == eAddressCollition.Equal)
                    {
                        if (EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.Transpose))
                        {
                            copiedMergedCells.Add(csem.Value, new ExcelAddress(
                                _destinationRange._fromRow + (adr.Start.Row - _sourceRange._fromRow),
                                _destinationRange._fromCol + (adr.Start.Column - _sourceRange._fromCol),
                                _destinationRange._fromRow + (adr.End.Column - _sourceRange._fromRow),
                                _destinationRange._fromCol + (adr.End.Row - _sourceRange._fromCol)));
                        }
                        else
                        {
                            copiedMergedCells.Add(csem.Value, new ExcelAddress(
                                _destinationRange._fromRow + (adr.Start.Row - _sourceRange._fromRow),
                                _destinationRange._fromCol + (adr.Start.Column - _sourceRange._fromCol),
                                _destinationRange._fromRow + (adr.End.Row - _sourceRange._fromRow),
                                _destinationRange._fromCol + (adr.End.Column - _sourceRange._fromCol)));
                        }
                    }
                    else
                    {
                        //Partial merge of the address ignore.
                        copiedMergedCells.Add(csem.Value, null);
                    }
                }
            }

            return copiedMergedCells;
        }

        private void CopyMergedCells(Dictionary<int, ExcelAddress> copiedMergedCells)
        {
            //Add merged cells
            foreach (var m in copiedMergedCells.Values)
            {
                if (m != null)
                {
                    _destinationRange._worksheet.MergedCells.Add(m, true);
                }
            }
        }

        private void CopyFullRow()
        {
            if (_sourceRange._fromRow == 1 && _sourceRange._toRow == ExcelPackage.MaxRows)
            {
                for (int col = 0; col < _sourceRange.Columns; col++)
                {
                    _destinationRange.Worksheet.Column(_destinationRange.Start.Column + col).OutlineLevel = _sourceRange.Worksheet.Column(_sourceRange._fromCol + col).OutlineLevel;
                }
            }
        }

        private void CopyFullColumn()
        {
            if (_sourceRange._fromCol == 1 && _sourceRange._toCol == ExcelPackage.MaxColumns)
            {
                for (int row = 0; row < _sourceRange.Rows; row++)
                {
                    _destinationRange.Worksheet.Row(_destinationRange.Start.Row + row).OutlineLevel = _sourceRange.Worksheet.Row(_sourceRange._fromRow + row).OutlineLevel;
                }
            }
        }
    }
}
