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
using EPPlus.Export.Pdf.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
{
    internal class PdfToUnicodeCMap : PdfObject
    {
        private readonly Dictionary<ushort, string> CharacterMappings;
        private readonly int CodeSpaceMin;
        private readonly int CodeSpaceMax;
        private readonly int BytesPerCode;

        /// <summary>
        /// Creates a ToUnicode CMap for mapping character codes to Unicode values
        /// </summary>
        /// <param name="objectNumber">PDF object number</param>
        /// <param name="characterMappings">Dictionary mapping character codes to Unicode strings (UTF-16BE hex format)</param>
        /// <param name="codeSpaceMin">Minimum character code value (e.g., 0 for simple fonts)</param>
        /// <param name="codeSpaceMax">Maximum character code value (e.g., 255 for simple fonts, 0xFFFF for CID fonts)</param>
        /// <param name="bytesPerCode">Number of bytes per character code (1 for simple fonts, 2 for CID fonts)</param>
        /// <param name="version">PDF object version</param>
        public PdfToUnicodeCMap(int objectNumber, Dictionary<ushort, string> characterMappings, int codeSpaceMin = 0, int codeSpaceMax = 65535, int bytesPerCode = 2, int version = 0) : base(objectNumber, version)
        {
            CharacterMappings = characterMappings ?? new Dictionary<ushort, string>();
            CodeSpaceMin = codeSpaceMin;
            CodeSpaceMax = codeSpaceMax;
            BytesPerCode = bytesPerCode;
        }

        internal override string RenderDictionary()
        {
            var cmapContent = GenerateCMapContent();
            var length = Encoding.UTF8.GetByteCount(cmapContent);
            var sb = new StringBuilder();
            sb.AppendLine(string.Format("<< /Length {0} >>", length.ToPdfStringF0()));
            sb.AppendLine("stream");
            sb.Append(cmapContent);
            sb.Append("\nendstream");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var cmapContent = GenerateCMapContent();
            var cmapBytes = Encoding.ASCII.GetBytes(cmapContent);
            var body = PdfFlate.Compress(cmapBytes);
            WriteAscii(bw, $"<< /Filter /FlateDecode /Length {body.Length.ToPdfStringF0()} >>\nstream\n");
            bw.Write(body);
            WriteAscii(bw, "\nendstream");
        }

        private string GenerateCMapContent()
        {
            var sb = new StringBuilder();
            sb.Append("/CIDInit /ProcSet findresource begin\n");
            sb.Append("12 dict begin\n");
            sb.Append("begincmap\n");
            sb.Append("/CIDSystemInfo\n");
            sb.Append("<< /Registry (Adobe)\n");
            sb.Append("   /Ordering (UCS)\n");
            sb.Append("   /Supplement 0\n");
            sb.Append(">> def\n");
            sb.Append("/CMapName /Adobe-Identity-UCS def\n");
            sb.Append("/CMapType 2 def\n");
            sb.Append("1 begincodespacerange\n");
            sb.AppendFormat($"<{FormatCode(CodeSpaceMin)}> <{FormatCode(CodeSpaceMax)}>\n");
            sb.Append("endcodespacerange\n");
            if (CharacterMappings.Count > 0)
            {
                GenerateCharacterMappings(sb);
            }
            sb.Append("endcmap\n");
            sb.Append("CMapName currentdict /CMap defineresource pop\n");
            sb.Append("end\n");
            sb.Append("end");
            return sb.ToString();
        }

        private void GenerateCharacterMappings(StringBuilder sb)
        {
            var sortedMappings = CharacterMappings.OrderBy(kvp => kvp.Key).ToList();
            var ranges = new List<CharacterRange>();
            var individualMappings = new List<CharacterMapping>();
            for (int i = 0; i < sortedMappings.Count; i++)
            {
                var current = sortedMappings[i];
                int unicodeValue;
                if (TryParseSimpleUnicode(current.Value, out unicodeValue))
                {
                    int rangeStart = current.Key;
                    int rangeEnd = current.Key;
                    int unicodeStart = unicodeValue;
                    // Check if we can extend this into a range
                    while (i + 1 < sortedMappings.Count)
                    {
                        var next = sortedMappings[i + 1];
                        int nextUnicode;
                        if (TryParseSimpleUnicode(next.Value, out nextUnicode) &&
                            next.Key == rangeEnd + 1 &&
                            nextUnicode == unicodeStart + (rangeEnd - rangeStart + 1))
                        {
                            rangeEnd = next.Key;
                            i++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (rangeEnd - rangeStart >= 2)
                    {
                        ranges.Add(new CharacterRange(rangeStart, rangeEnd, unicodeStart));
                    }
                    else
                    {
                        for (int j = rangeStart; j <= rangeEnd; j++)
                        {
                            var mapping = sortedMappings.FirstOrDefault(m => m.Key == j);
                            individualMappings.Add(new CharacterMapping(j, mapping.Value));
                        }
                    }
                }
                else
                {
                    individualMappings.Add(new CharacterMapping(current.Key, current.Value));
                }
            }
            // Output ranges
            if (ranges.Count > 0)
            {
                sb.AppendFormat($"{ranges.Count} beginbfrange");
                foreach (var range in ranges)
                {
                    sb.AppendFormat($"<{FormatCode(range.Start)}> <{FormatCode(range.End)}> <{FormatUnicode(range.UnicodeStart)}>\n");
                }
                sb.Append("endbfrange\n");
            }
            // Output individual character mappings
            if (individualMappings.Count > 0)
            {
                const int batchSize = 100;
                for (int i = 0; i < individualMappings.Count; i += batchSize)
                {
                    var batch = individualMappings.Skip(i).Take(batchSize).ToList();
                    sb.AppendFormat($"{batch.Count} beginbfchar\n");
                    foreach (var mapping in batch)
                    {
                        string hex = ((int)mapping.Unicode[0]).ToString("X4");
                        sb.AppendFormat($"<{FormatCode(mapping.Code)}> <{hex}>\n");
                    }
                    sb.Append("endbfchar\n");
                }
            }
        }

        private bool TryParseSimpleUnicode(string hexString, out int unicodeValue)
        {
            unicodeValue = 0;
            if (hexString.Length == 4 && int.TryParse(hexString, System.Globalization.NumberStyles.HexNumber, null, out unicodeValue))
            {
                return true;
            }
            return false;
        }

        private string FormatCode(int code)
        {
            // Format character code as hex with appropriate byte length
            if (BytesPerCode == 1)
            {
                return code.ToString("X2");
            }
            else if (BytesPerCode == 2)
            {
                return code.ToString("X4");
            }
            else if (BytesPerCode == 4)
            {
                return code.ToString("X8");
            }
            else
            {
                return code.ToString("X4");
            }
        }

        private string FormatUnicode(int unicodeValue)
        {
            if (unicodeValue <= 0xFFFF)
            {
                return unicodeValue.ToString("X4");
            }
            else
            {
                unicodeValue -= 0x10000;
                int highSurrogate = 0xD800 + (unicodeValue >> 10);
                int lowSurrogate = 0xDC00 + (unicodeValue & 0x3FF);
                return string.Format("{0:X4}{1:X4}", highSurrogate, lowSurrogate);
            }
        }

        private class CharacterRange
        {
            public int Start { get; private set; }
            public int End { get; private set; }
            public int UnicodeStart { get; private set; }

            public CharacterRange(int start, int end, int unicodeStart)
            {
                Start = start;
                End = end;
                UnicodeStart = unicodeStart;
            }
        }

        private class CharacterMapping
        {
            public int Code { get; private set; }
            public string Unicode { get; private set; }

            public CharacterMapping(int code, string unicode)
            {
                Code = code;
                Unicode = unicode;
            }
        }
    }
}
