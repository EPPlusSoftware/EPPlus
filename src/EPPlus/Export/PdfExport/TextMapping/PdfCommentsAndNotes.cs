/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.ThreadedComments;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport
{
    internal class PdfCommentsAndNotes
    {
        public ExcelComment Comment;
        public ExcelThreadedCommentThread ThreadedComment;
        public static bool HasThreadedComment = false;

        public PdfCommentsAndNotes(ExcelComment comment)
        {
            Comment = comment;
        }

        public PdfCommentsAndNotes(ExcelThreadedCommentThread tComment)
        {
            ThreadedComment = tComment;
        }

        public static ExcelWorksheet CreateCommentAndNotesPages(Dictionary<string, PdfCommentsAndNotes> CommentsAndNotesCollections, ExcelWorksheet ws)
        {
            var ns = ws.Workbook.Styles.GetNormalStyle();
            var tempWS = ws.Workbook.Worksheets.Add("TemporaryWorksheetForCommentsInPdfExporterForEPPlus");
            int row = 1;
            int col = 1;
            tempWS.Column(col).Width = 10d;
            tempWS.Column(col + 1).Width = 75d;
            foreach (var commentNote in CommentsAndNotesCollections)
            {
                AddText(tempWS, row, col, "Cell:", true, ExcelHorizontalAlignment.Right, ExcelVerticalAlignment.Bottom, ns);
                AddText(tempWS, row, col + 1, commentNote.Key, false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Bottom, ns);
                row++;
                if (commentNote.Value.ThreadedComment != null)
                {
                    var CommentReply = "Comment:";
                    foreach (var comment in commentNote.Value.ThreadedComment.Comments)
                    {
                        AddText(tempWS, row, col, CommentReply, true, ExcelHorizontalAlignment.Right, ExcelVerticalAlignment.Bottom, ns);
                        AddText(tempWS, row, col + 1, comment.Author.DisplayName, false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Bottom, ns);
                        row++;
                        AddText(tempWS, row, col + 1, comment.Text, false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Top, ns);
                        row++;
                        AddText(tempWS, row, col + 1, comment.DateCreated.ToString("yyyy-MM-dd HH:mm"), false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Bottom, ns);
                        row++;
                        CommentReply = "Reply:";
                    }
                }
                else if (commentNote.Value.Comment != null)
                {
                    var note = commentNote.Value.Comment;
                    AddText(tempWS, row, col, "Note:", true, ExcelHorizontalAlignment.Right, ExcelVerticalAlignment.Bottom, ns);
                    AddText(tempWS, row, col + 1, note.Author, false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Bottom, ns);
                    row++;
                    AddText(tempWS, row, col + 1, note.RichText, false, ExcelHorizontalAlignment.Left, ExcelVerticalAlignment.Bottom, ns);
                    row++;
                }
                row++;
            }
            return tempWS;
        }
        private static void AddText(ExcelWorksheet ws, int row, int col, string text, bool bold, ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment, ExcelNamedStyleXml ns)
        {
            var cell = ws.Cells[row, col];
            cell.RichText.Clear();
            var rt = cell.RichText.Add(text);
            rt.Bold = bold;
            rt.FontName = ns.Style.Font.Name;
            rt.Family = ns.Style.Font.Family;
            rt.Size = ns.Style.Font.Size;
            cell.Style.HorizontalAlignment = horizontalAlignment;
            cell.Style.VerticalAlignment = verticalAlignment;
            cell.Style.WrapText = true;
        }

        private static void AddText(ExcelWorksheet ws, int row, int col, ExcelRichTextCollection RichText, bool bold, ExcelHorizontalAlignment horizontalAlignment, ExcelVerticalAlignment verticalAlignment, ExcelNamedStyleXml ns)
        {
            var cell = ws.Cells[row, col];
            cell.RichText.Clear();
            for (int i = 1; i < RichText.Count; i++)
            {
                var rt = RichText[i];
                var trimmedText = rt.Text.Trim();
                var r = cell.RichText.Add(trimmedText);
                r.Bold = rt.Bold;
                r.Italic = rt.Italic;
                r.Color = rt.Color;
                r.ColorSettings = rt.ColorSettings;
                r.Strike = rt.Strike;
                r.UnderLine = rt.UnderLine;
                r.UnderLineType = rt.UnderLineType;
                r.FontName = rt.FontName;
                r.Family = rt.Family;
                r.Size = rt.Size;
            }
            cell.Style.HorizontalAlignment = horizontalAlignment;
            cell.Style.VerticalAlignment = verticalAlignment;
            cell.Style.WrapText = true;
        }
    }
}
