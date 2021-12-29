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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
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
        private readonly ExcelRangeBase _destination;
        private readonly ExcelRangeCopyOptionFlags _copyOptions;
        Dictionary<ulong, CopiedCell> _copiedCells=new Dictionary<ulong, CopiedCell>();
        internal RangeCopyHelper(ExcelRangeBase sourceRange, ExcelRangeBase destination, ExcelRangeCopyOptionFlags copyOptions)
        {
            _sourceRange = sourceRange;
            _destination = destination;
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

            CopyFullColumn();
            CopyFullRow();
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
                        if (_sourceRange._worksheet == _destination._worksheet)
                        {
                            dv.SetAddress(dv.Address + "," + newAddress);
                        }
                        else
                        {
                            _destination._worksheet.DataValidations.AddCopyOfDataValidation(new ExcelAddressBase(newAddress).AddressSpaceSeparated, dv);
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
                    if (_sourceRange._worksheet == _destination._worksheet)
                    {
                        cf.Address = new ExcelAddress(cf.Address + "," + newAddress);
                    }
                    else
                    {
                        _destination._worksheet.ConditionalFormatting.AddFromXml(new ExcelAddress(newAddress), cf.PivotTable, cf.Node.OuterXml);
                        if (cf.Style.HasValue)
                        {
                            var destRule = ((ExcelConditionalFormattingRule)_destination._worksheet.ConditionalFormatting[_destination._worksheet.ConditionalFormatting.Count - 1]);
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
                var fr = Math.Min(Math.Max(_destination._fromRow + rowOffset, 1), ExcelPackage.MaxRows);
                var fc = Math.Min(Math.Max(_destination._fromCol + colOffset, 1), ExcelPackage.MaxColumns);
                address = new ExcelAddressBase(fr, fc, Math.Min(fr + address.Rows-1, ExcelPackage.MaxRows), Math.Min(fc + address.Columns-1, ExcelPackage.MaxColumns));
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
            bool sameWorkbook = _destination._worksheet.Workbook == _sourceRange._worksheet.Workbook;

            AddValuesFormulasAndStyles(worksheet, includeStyles, styleCashe, sameWorkbook);

            if (includeComments)
            {
                AddComments(worksheet);
            }

            if (includeThreadedComments)
            {
                AddThreadedComments(worksheet);
            }
        }

        private void AddValuesFormulasAndStyles(ExcelWorksheet worksheet, bool includeStyles, Dictionary<int, int> styleCashe, bool sameWorkbook)
        {
            int styleId = 0;
            object o = null;
            byte flag = 0;
            Uri hl = null;

            var includeValues = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues);
            var includeFormulas = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas);
            var includeHyperlinks = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks);
            var cse = new CellStoreEnumerator<ExcelValue>(worksheet._values,  _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cell = new CopiedCell
                {
                    Row = _destination._fromRow + (row - _sourceRange._fromRow),
                    Column = _destination._fromCol + (col - _sourceRange._fromCol),
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
                            _destination._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
                        }
                    }
                    else
                    {
                        cell.Formula = o;
                    }
                }

                if (includeStyles && worksheet.ExistsStyleInner(row, col, ref styleId))
                {
                    if (sameWorkbook)
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
                            styleId = _destination._worksheet.Workbook.Styles.CloneStyle(_sourceRange._worksheet.Workbook.Styles, styleId);
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
                        Row = _destination._fromRow + (row - _sourceRange._fromRow),
                        Column = _destination._fromCol + (col - _sourceRange._fromCol),
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
                        Row = _destination._fromRow + (row - _sourceRange._fromRow),
                        Column = _destination._fromCol + (col - _sourceRange._fromCol),
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
                    _destination._worksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID ?? 0);
                }
                else if(EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
                {
                    _destination._worksheet.SetValueInner(cell.Row, cell.Column, cell.Value);
                }
                else
                {
                    _destination._worksheet.SetValueStyleIdInner(cell.Row, cell.Column, cell.Value, cell.StyleID ?? 0);
                }

                if ((EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas) && EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues)) &&
                    cell.Formula != null)
                {
                    cell.Formula = ExcelRangeBase.UpdateFormulaReferences(cell.Formula.ToString(), _destination._fromRow - fromRow, _destination._fromCol - fromCol, 0, 0, _destination.WorkSheetName, _destination.WorkSheetName, true, true);
                    _destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }

                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks) && 
                    cell.HyperLink != null)
                {
                    _destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
                }

                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments) && 
                    cell.ThreadedComment != null)
                {
                    var differentPackages = _destination._workbook != _sourceRange._workbook;
                    var tc = _destination.Worksheet.Cells[cell.Row, cell.Column].AddThreadedComment();
                    foreach (var c in cell.ThreadedComment.Comments)
                    {
                        if(differentPackages && _destination._workbook.ThreadedCommentPersons[c.PersonId]==null)
                        {
                            var p = _sourceRange._workbook.ThreadedCommentPersons[c.PersonId];
                            _destination._workbook.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                        }
                        tc.AddCommentFromXml((XmlElement)c.TopNode);
                    }
                }
                else if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments) && 
                    cell.Comment != null)
                {
                    CopyComment(_destination, cell);
                }

                if (cell.Flag != 0)
                {
                    _destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
                }

                if(EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas | ExcelRangeCopyOptionFlags.ExcludeValues) &&
                    cell.MetaData.cm > 0 || cell.MetaData.vm > 0)
                {
                    _destination._worksheet._metadataStore.SetValue(cell.Row, cell.Column, cell.MetaData);
                }
            }
        }
        private static void CopyComment(ExcelRangeBase destination, CopiedCell cell)
        {
            var c = destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment.Text, cell.Comment.Author);
            var offsetCol = c.Column - cell.Comment.Column;
            var offsetRow = c.Row - cell.Comment.Row;
            XmlHelper.CopyElement((XmlElement)cell.Comment.TopNode, (XmlElement)c.TopNode, new string[] { "id" });

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
            c.RichText = new Style.ExcelRichTextCollection(c._commentHelper.NameSpaceManager, c._commentHelper.GetNode("d:text"));
            //Add relation to image used for filling the comment
            if(cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Frame ||
              cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Tile ||
              cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Pattern)
            {
                var img = cell.Comment.Fill.PatternPictureSettings.Image;                
                c.Fill.PatternPictureSettings.Image = img;
            }
        }

        private void ClearDestination()
        {
            //Clear all existing cells; 
            int rows = _sourceRange._toRow - _sourceRange._fromRow + 1,
                cols = _sourceRange._toCol - _sourceRange._fromCol + 1;

            _destination._worksheet.MergedCells.Clear(new ExcelAddressBase(_destination._fromRow, _destination._fromCol, _destination._fromRow + rows - 1, _destination._fromCol + cols - 1));

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues) && EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
            {
                _destination._worksheet._values.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            }
            _destination._worksheet._formulas.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._metadataStore.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._hyperLinks.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._flags.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._commentsStore.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._threadedCommentsStore.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
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
                        copiedMergedCells.Add(csem.Value, new ExcelAddress(
                            _destination._fromRow + (adr.Start.Row - _sourceRange._fromRow),
                            _destination._fromCol + (adr.Start.Column - _sourceRange._fromCol),
                            _destination._fromRow + (adr.End.Row - _sourceRange._fromRow),
                            _destination._fromCol + (adr.End.Column - _sourceRange._fromCol)));
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
                    _destination._worksheet.MergedCells.Add(m, true);
                }
            }
        }

        private void CopyFullRow()
        {
            if (_sourceRange._fromRow == 1 && _sourceRange._toRow == ExcelPackage.MaxRows)
            {
                for (int col = 0; col < _sourceRange.Columns; col++)
                {
                    _destination.Worksheet.Column(_destination.Start.Column + col).OutlineLevel = _sourceRange.Worksheet.Column(_sourceRange._fromCol + col).OutlineLevel;
                }
            }
        }

        private void CopyFullColumn()
        {
            if (_sourceRange._fromCol == 1 && _sourceRange._toCol == ExcelPackage.MaxColumns)
            {
                for (int row = 0; row < _sourceRange.Rows; row++)
                {
                    _destination.Worksheet.Row(_destination.Start.Row + row).OutlineLevel = _sourceRange.Worksheet.Row(_sourceRange._fromRow + row).OutlineLevel;
                }
            }
        }
    }
}
