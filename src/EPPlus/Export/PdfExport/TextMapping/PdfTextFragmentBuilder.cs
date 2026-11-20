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
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfTextFragmentBuilder
    {
        internal static List<TextFragment> GetTextFragments(PdfRenderSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
            bool dxfBold, dxfItalic, dxfStrike, dxfUnderline;
            ExcelUnderLineType dxfUnderLineType;
            System.Drawing.Color? dxfColor;
            ReadDxfFontOverrides(cellStyle, out dxfBold, out dxfItalic, out dxfStrike, out dxfUnderline, out dxfUnderLineType, out dxfColor);
            string forcedText = ResolveErrorText(pageSettings, cell);
            if (cell.IsRichText)
            {
                return GetTextFragmentsFromRichText(pageSettings, dictionaries, cell.RichText, forcedText, dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType, dxfColor);
            }
            return GetTextFragmentsFromCellStyle(pageSettings, dictionaries, cell, forcedText, dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType, dxfColor);
        }

        private static List<TextFragment> GetTextFragmentsFromRichText(PdfRenderSettings pageSettings, PdfDictionaries dictionaries, ExcelRichTextCollection richText, string forcedText, bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType, System.Drawing.Color? dxfColor)
        {
            var textFragments = new List<TextFragment>(richText.Count);
            for (int i = 0; i < richText.Count; i++)
            {
                var rt = richText[i];
                var textFrag = new TextFragment();
                textFrag.Font = new RichTextFormatSimple();
                textFrag.Text = forcedText == null ? rt.Text : forcedText;
                textFrag.Font.Family = rt.FontName;
                textFrag.Font.Size = rt.Size;
                textFrag.RichTextOptions.Bold = rt.Bold || dxfBold;
                textFrag.RichTextOptions.Italic = rt.Italic || dxfItalic;
                textFrag.RichTextOptions.UnderlineType = MapUnderlineType(rt.UnderLineType);
                textFrag.RichTextOptions.StrikeType = (rt.Strike || dxfStrike) ? 2 : 1;
                textFrag.RichTextOptions.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                textFrag.RichTextOptions.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                textFrag.RichTextOptions.FontColor = dxfColor ?? rt.Color;
                textFrag.Font.SubFamily = ComputeFontStyle(textFrag);
                textFragments.Add(textFrag);
                dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            }
            return textFragments;
        }

        private static List<TextFragment> GetTextFragmentsFromCellStyle(PdfRenderSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, string forcedText, bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType, System.Drawing.Color? dxfColor)
        {
            var font = cell.Style.Font;
            var textFragments = new List<TextFragment>(1);
            var textFrag = new TextFragment();
            textFrag.Font = new RichTextFormatSimple();
            textFrag.Text = forcedText == null ? cell.Text : forcedText;
            textFrag.Font.Family = font.Name;
            textFrag.Font.Size = font.Size;
            textFrag.RichTextOptions.Bold = font.Bold || dxfBold;
            textFrag.RichTextOptions.Italic = font.Italic || dxfItalic;
            ExcelUnderLineType underLineType;
            if (font.UnderLine)
            {
                underLineType = font.UnderLineType == ExcelUnderLineType.None ? ExcelUnderLineType.Single : font.UnderLineType;
            }
            else if (dxfUnderline)
            {
                underLineType = dxfUnderLineType;
            }
            else
            {
                underLineType = ExcelUnderLineType.None;
            }
            textFrag.RichTextOptions.UnderlineType = MapUnderlineType(underLineType);
            textFrag.RichTextOptions.StrikeType = (font.Strike || dxfStrike) ? 2 : 1;
            textFrag.RichTextOptions.SuperScript = font.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
            textFrag.RichTextOptions.SubScript = font.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
            bool cellColorWins = !HasConditionalFontColor(cell) && CellFontColorIsUserSet(cell);
            textFrag.RichTextOptions.FontColor = cellColorWins ? font.Color.ToColor() : (dxfColor ?? font.Color.ToColor());
            textFrag.Font.SubFamily = ComputeFontStyle(textFrag);
            textFragments.Add(textFrag);
            dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            return textFragments;
        }

        private static void ReadDxfFontOverrides(PdfCellStyle cellStyle, out bool bold, out bool italic, out bool strike, out bool underline, out ExcelUnderLineType underLineType, out System.Drawing.Color? color)
        {
            bold = false; italic = false; strike = false; underline = false;
            underLineType = ExcelUnderLineType.None; color = null;
            if (cellStyle == null) return;
            var elem = cellStyle.dxfFont;
            var ov = cellStyle.dxfFontOverride;
            bold = (elem?.Bold == true) || (ov?.Bold == true);
            italic = (elem?.Italic == true) || (ov?.Italic == true);
            strike = (elem?.Strike == true) || (ov?.Strike == true);
            var un = ov?.Underline ?? elem?.Underline;
            underline = un != null;
            underLineType = un != null ? (ExcelUnderLineType)un : ExcelUnderLineType.None;
            var cl = (ov?.Color != null && ov.Color.HasValue) ? ov.Color : (elem?.Color != null && elem.Color.HasValue) ? elem.Color : null;
            if (cl != null && cl.HasValue)
            {
                if (cl.Auto == true)
                {
                    color = System.Drawing.Color.Black;
                }
                else
                {
                    var baseColor = cl.GetColorAsColor();
                    if (!baseColor.IsEmpty)
                    {
                        var t = cl.Tint ?? 0d;
                        color = t != 0d ? PdfThemeColorMath.ApplyThemeTint(baseColor, t) : baseColor;
                    }
                }
            }
        }

        private static bool CellFontColorIsUserSet(ExcelRangeBase cell)
        {
            try
            {
                var cellColor = cell.Style.Font.Color;
                var normal = cell.Worksheet.Workbook.Styles.GetNormalStyle();
                var normalColor = normal?.Style?.Font?.Color;
                var c = cellColor.ToColor();
                if (normalColor == null)
                {
                    return c.R != 0 || c.G != 0 || c.B != 0;
                }
                var n = normalColor.ToColor();
                return c.R != n.R || c.G != n.G || c.B != n.B;
            }
            catch
            {
                return false;
            }
        }

        private static bool HasConditionalFontColor(ExcelRangeBase cell)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf == null || cf.Count == 0) return false;
            foreach (var rule in cf.OrderBy(r => r.Priority))
            {
                if (!rule.ShouldApplyToCell(cell)) { if (rule.StopIfTrue) break; continue; }
                if (rule.Style?.Font != null && rule.Style.Font.HasValue) return true;
                if (rule.StopIfTrue) break;
            }
            return false;
        }

        private static int MapUnderlineType(ExcelUnderLineType type)
        {
            // 12 = none, 13 = single, 4 = double (matches existing rendering code; accounting not supported).
            if (type == ExcelUnderLineType.Single) return 13;
            if (type == ExcelUnderLineType.Double) return 4;
            return 12;
        }

        internal static FontSubFamily ComputeFontStyle(TextFragment textFrag)
        {
            if (textFrag.RichTextOptions.Bold && textFrag.RichTextOptions.Italic) return FontSubFamily.BoldItalic;
            if (textFrag.RichTextOptions.Bold) return FontSubFamily.Bold;
            if (textFrag.RichTextOptions.Italic) return FontSubFamily.Italic;
            return FontSubFamily.Regular;
        }

        private static string ResolveErrorText(PdfRenderSettings pageSettings, ExcelRangeBase cell)
        {
            if (!ExcelErrorValue.IsErrorValue(cell.Text)) return null;
            switch (pageSettings.CellErrors)
            {
                case CellErrors.Blank: return "";
                case CellErrors.Dashed: return "--";
                case CellErrors.NA: return "#N/A";
                case CellErrors.Displayed:
                default: return null;
            }
        }
    }
}
