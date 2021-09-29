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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core
{
    internal static class RangeCopyHelper
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
            internal Byte Flag { get; set; }
            internal ExcelWorksheet.MetaDataReference MetaData{ get; set; }
    }

        internal static void Copy(ExcelRangeBase sourceRange, ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
        {
            var copiedValue = GetCopiedValues(sourceRange, Destination, excelRangeCopyOptionFlags);
            var copiedMergedCells = GetCopiedMergedCells(sourceRange, Destination);

            //Clear all existing cells; 
            int rows = sourceRange._toRow - sourceRange._fromRow + 1,
                cols = sourceRange._toCol - sourceRange._fromCol + 1;
            ClearDestination(Destination, rows, cols);

            CopyValues(Destination, sourceRange, copiedValue);

            CopyMergedCells(Destination, copiedMergedCells);
            CopyFullColumn(sourceRange, Destination);
            CopyFullRow(sourceRange, Destination);

        }

        private static List<CopiedCell> GetCopiedValues(ExcelRangeBase sourceRange, ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
        {
            var worksheet = sourceRange._worksheet;
            var toRow = sourceRange._toRow;
            var toCol = sourceRange._toCol;
            var fromRow = sourceRange._fromRow;
            var fromCol = sourceRange._fromCol;

            int i = 0;
            object o = null;
            byte flag = 0;
            Uri hl = null;

            var excludeFormulas = (excelRangeCopyOptionFlags ?? 0 & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;

            var copiedValue = new List<CopiedCell>();
            ExcelStyles sourceStyles = worksheet.Workbook.Styles, styles = Destination._worksheet.Workbook.Styles;
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            bool sameWorkbook = Destination._worksheet.Workbook == sourceRange._worksheet.Workbook;

            var cse = new CellStoreEnumerator<ExcelValue>(worksheet._values, fromRow, fromCol, toRow, toCol);
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cell = new CopiedCell
                {
                    Row = Destination._fromRow + (row - fromRow),
                    Column = Destination._fromCol + (col - fromCol),
                    Value = cse.Value._value
                };

                if (!excludeFormulas && worksheet._formulas.Exists(row, col, ref o))
                {
                    if (o is int)
                    {
                        cell.Formula = worksheet.GetFormula(cse.Row, cse.Column);
                        if (worksheet._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.ArrayFormula))
                        {
                            Destination._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
                        }
                    }
                    else
                    {
                        cell.Formula = o;
                    }
                }
                if (worksheet.ExistsStyleInner(row, col, ref i))
                {
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
                            i = styles.CloneStyle(sourceStyles, i);
                            styleCashe.Add(oldStyleID, i);
                        }
                        cell.StyleID = i;
                    }
                }

                var md = new ExcelWorksheet.MetaDataReference();
                if (worksheet._metadataStore.Exists(row, col, ref md))
                {
                    cell.MetaData=md;
                }

                if (worksheet._hyperLinks.Exists(row, col, ref hl))
                {
                    cell.HyperLink = hl;
                }

                // Will just be null if no comment exists.
                cell.Comment = worksheet.Cells[cse.Row, cse.Column].Comment;
                cell.ThreadedComment = worksheet.Cells[cse.Row, cse.Column].ThreadedComment;
                if (worksheet._flags.Exists(row, col, ref flag))
                {
                    cell.Flag = flag;
                }
                copiedValue.Add(cell);
            }

            //Copy styles with no cell value
            var cses = new CellStoreEnumerator<ExcelValue>(worksheet._values, fromRow, fromCol, toRow, toCol);
            while (cses.Next())
            {
                if (!worksheet.ExistsValueInner(cses.Row, cses.Column))
                {
                    var row = Destination._fromRow + (cses.Row - fromRow);
                    var col = Destination._fromCol + (cses.Column - fromCol);
                    var cell = new CopiedCell
                    {
                        Row = row,
                        Column = col,
                        Value = null
                    };

                    i = cses.Value._styleId;
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
                            i = styles.CloneStyle(sourceStyles, i);
                            styleCashe.Add(oldStyleID, i);
                        }
                        cell.StyleID = i;
                    }
                    copiedValue.Add(cell);
                }
            }

            return copiedValue;
        }

        private static void CopyValues(ExcelRangeBase destination, ExcelRangeBase source, List<CopiedCell> copiedValue)
        {
            int fromRow = source._fromRow;
            int fromCol = source._fromCol;
            foreach (var cell in copiedValue)
            {
                destination._worksheet.SetValueStyleIdInner(cell.Row, cell.Column, cell.Value, cell.StyleID??0);

                if (cell.Formula != null)
                {
                    cell.Formula = ExcelRangeBase.UpdateFormulaReferences(cell.Formula.ToString(), destination._fromRow - fromRow, destination._fromCol - fromCol, 0, 0, destination.WorkSheetName, destination.WorkSheetName, true, true);
                    destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }
                if (cell.HyperLink != null)
                {
                    destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
                }

                if (cell.ThreadedComment != null)
                {
                    var differentPackages = destination._workbook != source._workbook;
                    var tc = destination.Worksheet.Cells[cell.Row, cell.Column].AddThreadedComment();
                    foreach (var c in cell.ThreadedComment.Comments)
                    {
                        if(differentPackages && destination._workbook.ThreadedCommentPersons[c.PersonId]==null)
                        {
                            var p = source._workbook.ThreadedCommentPersons[c.PersonId];
                            destination._workbook.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                        }
                        tc.AddCommentFromXml((XmlElement)c.TopNode);
                    }
                }
                else if (cell.Comment != null)
                {
                    CopyComment(destination, cell);
                }

                if (cell.Flag != 0)
                {
                    destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
                }

                if(cell.MetaData.cm > 0 || cell.MetaData.vm > 0)
                {
                    destination._worksheet._metadataStore.SetValue(cell.Row, cell.Column, cell.MetaData);
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

        private static void ClearDestination(ExcelRangeBase Destination, int rows, int cols)
        {
            Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow + rows - 1, Destination._fromCol + cols - 1));

            Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._threadedCommentsStore.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
            Destination._worksheet._metadataStore.Clear(Destination._fromRow, Destination._fromCol, rows, cols);
        }

        private static Dictionary<int, ExcelAddress> GetCopiedMergedCells(ExcelRangeBase sourceRange, ExcelRangeBase Destination)
        {
            var toRow = sourceRange._toRow;
            var toCol = sourceRange._toCol;
            var fromRow = sourceRange._fromRow;
            var fromCol = sourceRange._fromCol;

            var worksheet = sourceRange._worksheet;
            var copiedMergedCells = new Dictionary<int, ExcelAddress>();
            //Merged cells
            var csem = new CellStoreEnumerator<int>(worksheet.MergedCells._cells, fromRow, fromCol, toRow, toCol);
            while (csem.Next())
            {
                if (!copiedMergedCells.ContainsKey(csem.Value))
                {
                    var adr = new ExcelAddress(worksheet.Name, worksheet.MergedCells._list[csem.Value]);
                    var collideResult = sourceRange.Collide(adr);
                    if (collideResult == eAddressCollition.Inside || collideResult == eAddressCollition.Equal)
                    {
                        copiedMergedCells.Add(csem.Value, new ExcelAddress(
                            Destination._fromRow + (adr.Start.Row - fromRow),
                            Destination._fromCol + (adr.Start.Column - fromCol),
                            Destination._fromRow + (adr.End.Row - fromRow),
                            Destination._fromCol + (adr.End.Column - fromCol)));
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

        private static void CopyMergedCells(ExcelRangeBase Destination, Dictionary<int, ExcelAddress> copiedMergedCells)
        {
            //Add merged cells
            foreach (var m in copiedMergedCells.Values)
            {
                if (m != null)
                {
                    Destination._worksheet.MergedCells.Add(m, true);
                }
            }
        }

        private static void CopyFullRow(ExcelRangeBase sourceRange, ExcelRangeBase Destination)
        {
            if (sourceRange._fromRow == 1 && sourceRange._toRow == ExcelPackage.MaxRows)
            {
                for (int col = 0; col < sourceRange.Columns; col++)
                {
                    Destination.Worksheet.Column(Destination.Start.Column + col).OutlineLevel = sourceRange.Worksheet.Column(sourceRange._fromCol + col).OutlineLevel;
                }
            }
        }

        private static void CopyFullColumn(ExcelRangeBase sourceRange, ExcelRangeBase Destination)
        {
            if (sourceRange._fromCol == 1 && sourceRange._toCol == ExcelPackage.MaxColumns)
            {
                for (int row = 0; row < sourceRange.Rows; row++)
                {
                    Destination.Worksheet.Row(Destination.Start.Row + row).OutlineLevel = sourceRange.Worksheet.Row(sourceRange._fromRow + row).OutlineLevel;
                }
            }
        }
    }
}
