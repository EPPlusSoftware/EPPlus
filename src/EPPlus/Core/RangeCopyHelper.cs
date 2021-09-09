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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
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

            CopyValues();
            CopyConditionalFormatting();

            if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeMergedCells))
            {
                CopyMergedCells(copiedMergedCells);
            }

            CopyFullColumn();
            CopyFullRow();
        }

        private void CopyConditionalFormatting()
        {
            foreach(var cf in _sourceRange._worksheet.ConditionalFormatting)
            {
                if(cf.Address.Collide(_sourceRange)!=eAddressCollition.No)
                {
                    var address = cf.Address.Intersect(_sourceRange);
                    var rowOffset = address._fromRow - _sourceRange._fromRow;
                    var colOffset = address._fromCol - _sourceRange._fromCol;
                    address = new ExcelAddressBase(address._fromRow + rowOffset, address._fromCol + colOffset, address._toRow + rowOffset, address._toCol + colOffset);
                    var ruleXml = cf.Node.OuterXml;
                    
                }
            }
        }

        private void GetCopiedValues()
        {
            var worksheet = _sourceRange._worksheet;
            var toRow = _sourceRange._toRow;
            var toCol = _sourceRange._toCol;
            var fromRow = _sourceRange._fromRow;
            var fromCol = _sourceRange._fromCol;

            var includeValues = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulasAndValues);
            var includeStyles = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles);
            var includeComments = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments);
            var includeThreadedComments = EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments);

            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            bool sameWorkbook = _destination._worksheet.Workbook == _sourceRange._worksheet.Workbook;

            AddValuesFormulasAndStyles(worksheet, includeStyles, styleCashe, sameWorkbook);

            //if (includeStyles)
            //{
            //    AddStyles(worksheet, styleCashe, sameWorkbook);
            //}

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
                    Value = cse.Value._value
                };

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

        private void AddStyles(ExcelWorksheet worksheet, Dictionary<int, int> styleCashe, bool sameWorkbook)
        {
            //Copy styles with no cell value
            var cses = new CellStoreEnumerator<ExcelValue>(worksheet._values, _sourceRange._fromRow, _sourceRange._fromCol, _sourceRange._toRow, _sourceRange._toCol);
            while (cses.Next())
            {
                if (!worksheet.ExistsValueInner(cses.Row, cses.Column))
                {
                    var row = _destination._fromRow + (cses.Row - _sourceRange._fromRow);
                    var col = _destination._fromCol + (cses.Column - _sourceRange._fromRow);
                    var cell = new CopiedCell
                    {
                        Row = row,
                        Column = col,
                        Value = null
                    };

                    var i = cses.Value._styleId;
                    if (sameWorkbook)
                    {
                        cell.StyleID = i;
                    }
                    else
                    {
                        if (styleCashe.ContainsKey(i))
                        {
                            i = styleCashe[i];
                        }
                        else
                        {
                            var oldStyleID = i;
                            i = _destination._worksheet.Workbook.Styles.CloneStyle(_sourceRange._worksheet.Workbook.Styles, i);
                            styleCashe.Add(oldStyleID, i);
                        }
                        cell.StyleID = i;
                    }
                    _copiedCells.Add(ExcelCellBase.GetCellId(0, row, col), cell);
                }
            }
        }

        private void CopyValues()
        {
            int fromRow = _sourceRange._fromRow;
            int fromCol = _sourceRange._fromCol;
            foreach (var cell in _copiedCells.Values)
            {
                if (EnumUtil.HasFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulasAndValues) && 
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

                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulasAndValues | ExcelRangeCopyOptionFlags.ExcludeFormulas) &&
                    cell.Formula != null)
                {
                    cell.Formula = ExcelRangeBase.UpdateFormulaReferences(cell.Formula.ToString(), _destination._fromRow - fromRow, _destination._fromCol - fromCol, 0, 0, _destination.WorkSheetName, _destination.WorkSheetName, true, true);
                    _destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }
                if (EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulasAndValues | ExcelRangeCopyOptionFlags.ExcludeFormulas) && 
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

                if(EnumUtil.HasNotFlag(_copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas | ExcelRangeCopyOptionFlags.ExcludeFormulasAndValues) &&
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
            XmlHelper.CopyElement((XmlElement)cell.Comment.TopNode, (XmlElement)c.TopNode, new string[] { "Id" });

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
        }

        private void ClearDestination()
        {
            //Clear all existing cells; 
            int rows = _sourceRange._toRow - _sourceRange._fromRow + 1,
                cols = _sourceRange._toCol - _sourceRange._fromCol + 1;

            _destination._worksheet.MergedCells.Clear(new ExcelAddressBase(_destination._fromRow, _destination._fromCol, _destination._fromRow + rows - 1, _destination._fromCol + cols - 1));

            _destination._worksheet._values.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._formulas.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._hyperLinks.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._flags.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._commentsStore.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._threadedCommentsStore.Clear(_destination._fromRow, _destination._fromCol, rows, cols);
            _destination._worksheet._metadataStore.Clear (_destination._fromRow, _destination._fromCol, rows, cols);
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
