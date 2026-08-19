/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfPage : PdfObject
    {
        private readonly int parentObjectNumber;
        internal readonly List<int> contentObjectNumbers;
        PdfDictionaries dictionaries;
        internal PdfPageSize Size;

        public PdfPage(int objectNumber, int parentObjectNumber, List<int> contentObjectNumbers, PdfPageSize size, PdfDictionaries dictionaries, int version = 0)
            : base(objectNumber, version)
        {
            this.parentObjectNumber = parentObjectNumber;
            this.contentObjectNumbers = contentObjectNumbers;
            this.dictionaries = dictionaries;
            Size = size;
        }

        internal override string RenderDictionary()
        {
            var fontEntries = dictionaries.Fonts.Select(f => $"/{f.Value.Label} {f.Value.fontObjectNumber} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            var patternEntries = dictionaries.Patterns.Select(p => $"/{p.Value.Label} {p.Value.objectNumber} 0 R").ToArray();
            var patterns = string.Join(" ", patternEntries);
            var shadingEntries = dictionaries.Shadings.Select(s => $"/{s.Value.Label} {s.Value.objectNumber} 0 R").ToArray();
            var shadings = string.Join(" ", shadingEntries);
            var imageEntries = dictionaries.Images.Select(im => $"/{im.Value.Label} {im.Value.objectNumber} 0 R").ToArray();
            var images = string.Join(" ", imageEntries);
            var contentEntries = contentObjectNumbers.Select(con => $"{con} 0 R").ToArray();
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Page\n" +
                            $"   /Parent {parentObjectNumber} 0 R\n");
            bool hasFont = !string.IsNullOrEmpty(fonts);
            bool hasPattern = !string.IsNullOrEmpty(patterns);
            bool hasShading = !string.IsNullOrEmpty(shadings);
            bool hasImage = !string.IsNullOrEmpty(images);
            if (hasFont || hasPattern || hasShading)
            {
                sb.AppendFormat($"   /Resources <<\n");
                if (hasFont   ) sb.AppendFormat($"      /Font << {fonts} >>\n");
                if (hasPattern) sb.AppendFormat($"      /Pattern << {patterns} >>\n");
                if (hasShading) sb.AppendFormat($"      /Shading << {shadings} >>\n");
                if (hasImage) sb.AppendFormat($"      /XObject << {images} >>\n");
                sb.AppendFormat($"   >>\n");
            }
            sb.AppendFormat($"   /MediaBox [ 0 0 {Size.WidthPu.ToPdfString()} {Size.HeightPu.ToPdfString()} ]\n" +
                            $"   /Contents [ {string.Join(" ", contentEntries)} ] >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var fontEntries = dictionaries.Fonts.Select(f => $"/{f.Value.Label} {f.Value.fontObjectNumber} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            var patternEntries = dictionaries.Patterns.Select(p => $"/{p.Value.Label} {p.Value.objectNumber} 0 R").ToArray();
            var patterns = string.Join(" ", patternEntries);
            var shadingEntries = dictionaries.Shadings.Select(s => $"/{s.Value.Label} {s.Value.objectNumber} 0 R").ToArray();
            var shadings = string.Join(" ", shadingEntries);
            var imageEntries = dictionaries.Images.Select(im => $"/{im.Value.Label} {im.Value.objectNumber} 0 R").ToArray();
            var images = string.Join(" ", imageEntries);
            var contentEntries = contentObjectNumbers.Select(con => $"{con} 0 R").ToArray();
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Page\n" +
                           $"   /Parent {parentObjectNumber} 0 R\n");
            bool hasFont = !string.IsNullOrEmpty(fonts);
            bool hasPattern = !string.IsNullOrEmpty(patterns);
            bool hasShading = !string.IsNullOrEmpty(shadings);
            bool hasImage = !string.IsNullOrEmpty(images);
            if (hasFont || hasPattern || hasShading)
            {
                sb.AppendFormat($"   /Resources <<\n");
                if (hasFont   ) sb.AppendFormat($"      /Font << {fonts} >>\n");
                if (hasPattern) sb.AppendFormat($"      /Pattern << {patterns} >>\n");
                if (hasShading) sb.AppendFormat($"      /Shading << {shadings} >>\n");
                if (hasImage) sb.AppendFormat($"      /XObject << {images} >>\n");
                sb.AppendFormat($"   >>\n");
            }
            sb.AppendFormat($"   /MediaBox [ 0 0 {Size.WidthPu.ToPdfString()} {Size.HeightPu.ToPdfString()} ]\n" +
                            $"   /Contents [ {string.Join(" ", contentEntries)} ] >>");
            WriteAscii(bw, sb.ToString());
        }
    }
}
