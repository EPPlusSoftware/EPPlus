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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal class PdfAdditionalContent
    {
        internal static readonly Color TextColor = Color.Gray;
        internal const float FontSize = 7f;
        internal PdfCellBase content;

        public PdfAdditionalContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws)
        {
            if (string.IsNullOrEmpty(pageSettings.AdditionalContent))
            {
                return;
            }
            else
            {
                var ns = ws.Workbook.Styles.GetNormalStyle();
                var fragment = new TextFragment
                {
                    Font = new RichTextFormatSimple
                    {
                        Family = ns.Style.Font.Name,
                        Size = FontSize,
                        SubFamily = FontSubFamily.Regular,
                    },
                    Text = ComposeAdditionalContent(),
                };
                fragment.RichTextOptions.UnderlineType = 12;
                fragment.RichTextOptions.StrikeType = 1;
                fragment.RichTextOptions.FontColor = TextColor;
                content = new PdfCellBase
                {
                    TextFragments = new List<TextFragment> { fragment },
                    ContentAligmnet = new PdfCellAlignmentData
                    {
                        HorizontalAlignment = EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Left,
                        VerticalAlignment = EPPlus.Export.Pdf.Enums.ExcelVerticalAlignment.Bottom,
                    },
                };
                dictionaries.AddFont(pageSettings, fragment.Font.Family, fragment.Font.SubFamily, fragment.Text);
            }
        }

        private static string ComposeAdditionalContent()
        {
            byte[] d =
            {
                0xF3, 0xAE, 0x8C, 0x77, 0x03, 0x12, 0x25, 0xC6, 0xBF, 0xC9, 0xBC, 0x8F,
                0x3B, 0x5F, 0x21, 0x08, 0xF8, 0xC4, 0xA1, 0x91, 0x77, 0x12, 0x24, 0x03,
                0xE6, 0xC0, 0xAA, 0xCC, 0x4E, 0x7A, 0x19, 0x04, 0xF2, 0xD5, 0xE5, 0x91,
                0x6D, 0x46, 0x24, 0x12, 0x5F, 0xFF, 0x9D, 0xB2, 0x94, 0x74, 0x14, 0x3B,
                0x18, 0xFB, 0xD8, 0xB1, 0x81, 0x71, 0x58, 0x31, 0x03, 0xAE, 0xC1, 0xA5,
                0x88, 0x6F, 0x47, 0x3B, 0x02, 0xA6, 0xC3, 0xAB, 0x91, 0x22, 0x51, 0x25,
                0x2D, 0x0D, 0xF2, 0xD2, 0xBA, 0x96, 0x39, 0x57, 0x25, 0x56, 0xFB, 0xDB,
                0xBD, 0xDF, 0x72, 0x5F, 0x22, 0x03, 0xE8, 0xDE, 0xA8, 0x83, 0x68, 0x44,
                0x67, 0x13, 0xF6, 0xC1, 0xED, 0xC2, 0x55, 0x4F, 0x1F, 0x2C, 0x18, 0xF1,
                0xD4, 0xAC, 0x9C, 0x38, 0x43, 0x3E, 0x1C, 0xE7, 0x93, 0xBC, 0x9E, 0x64,
                0x46, 0x2D, 0x08, 0xA0, 0x8B, 0xBC, 0x80, 0x7B, 0x4E, 0x32, 0x45, 0xEC,
                0xD7, 0xB6, 0x91, 0x73, 0x25, 0x11, 0x72, 0x0B, 0xEC, 0xCD, 0xF7, 0x9D,
                0x67, 0x46, 0x39, 0x01, 0xE0, 0xC1, 0xBE, 0x96, 0x7B, 0x59, 0x2C, 0x1E,
                0xEE, 0x84, 0xAA, 0x87, 0x6A, 0x09, 0x65, 0x10, 0xEC, 0x82, 0xB1, 0x95,
                0x8D, 0x7D, 0x5C, 0x2F, 0x1E, 0xBA, 0xD8, 0xF8, 0x9B, 0x7F, 0x56, 0x31,
                0x1D, 0xE1, 0xD4, 0xFE,
            };
            const int a = 0xA7, m = 31;
            var buf = new char[d.Length];
            for (int i = 0; i < d.Length; i++)
                buf[i] = (char)(d[i] ^ ((a + i * m) & 0xFF));
            return new string(buf);
        }
    }
}
