using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.Linq;
using static EPPlus.Export.Pdf.PdfLayout.PdfCatalogLayout;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCommentsLayout
    {

        static double firstColumnWidth = 56.501953125d;
        static double secondColumnWidth = 418.8896484375d;

        public static void CreateCommentAndNotesPages(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws, PdfPagesLayout pagesLayout)
        {
            if (dictionaries.CommentsAndNotes.Count == 0) return;
            var ns = ws.Workbook.Styles.GetNormalStyle();
            var tempWS = ws.Workbook.Worksheets.Add("TemporaryWorksheetForCommentsInPdfExporterForEPPlus");
            int row = 1;
            int col = 1;
            tempWS.Column(col).Width = 10d;
            tempWS.Column(col + 1).Width = 75d;
            foreach (var commentNote in dictionaries.CommentsAndNotes)
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
            //get worksheet layout
            var setting = pageSettings;
            setting.CommentsAndNotes = CommentsAndNotes.None;
            var commentsLayout = new PdfWorksheetLayout(tempWS, setting, dictionaries);
            var s = commentsLayout.ToHierarchyString();
            setting.ShowHeadings = false;
            PdfCatalogLayout.CreatePageLayoutObjects(tempWS, setting, dictionaries, commentsLayout, pagesLayout, true);

            //populate pages
            var pages = pagesLayout.ChildObjects.Where(x => ((PdfPageLayout)x).isCommentsPage).ToArray();
            var pageData = new List<PdfCatalogLayout.PageData>();
            foreach (PdfPageLayout p in pages)
            {
                pageData.Add(new PdfCatalogLayout.PageData(p, p.ChildObjects[0].GetGlobalBoundingbox()));
            }
            pageData.Sort((a, b) => a.Bounds.Top.CompareTo(b.Bounds.Top));
            var transforms = new List<Transform>(commentsLayout.ChildObjects);
            foreach (var t in transforms)
            {
                var cellBounds = t.GetGlobalBoundingbox();
                // Pass 1: check if ANY page fully contains this transform. 
                // If so, we use only that page and skip all partial intersects.
                PageData fullIntersectPage = null;
                foreach (var pd in pageData)
                {
                    if (pd.Page.isCommentsPage) continue;
                    if (pd.Bounds.Top > cellBounds.Bottom) break;
                    if (pd.Bounds.Bottom < cellBounds.Top) continue;

                    if (Transform.IntersectsFully(pd.Bounds, cellBounds))
                    {
                        fullIntersectPage = pd;
                        break;
                    }
                }
                // Pass 2: assign to pages.
                foreach (var pd in pageData)
                {
                    if (pd.Page.isCommentsPage) continue;
                    if (pd.Bounds.Top > cellBounds.Bottom) break;
                    if (pd.Bounds.Bottom < cellBounds.Top) continue;

                    bool fullIntersect = fullIntersectPage != null && pd == fullIntersectPage;
                    bool partialIntersect = fullIntersectPage == null && !fullIntersect && Transform.Intersects(cellBounds, pd.Bounds);

                    if (!fullIntersect && !partialIntersect) continue;

                    var page = pd.Page;

                    if (t is PdfCellContentLayout cellContent)
                        page.AddChild(cellContent);
                }
            }
            ws.Workbook.Worksheets.Delete("TemporaryWorksheetForCommentsInPdfExporterForEPPlus");
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
            foreach (var rt in RichText)
            {
                var r = cell.RichText.Add(rt.Text);
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

        private static PdfPageLayout CreatePage(PdfPageSettings pageSettings, int pageNumber)
        {
            PdfPageLayout page = new PdfPageLayout(0, 0, pageSettings.PageSize.WidthPu, pageSettings.PageSize.HeightPu);
            page.Name = "CommentsAndNotes " + pageNumber;
            page.isCommentsPage = true;
            PdfContentLayout content = new PdfContentLayout(0, 0, pageSettings.ContentBounds);
            content.Parent = page;
            return page;
        }
    }
}
