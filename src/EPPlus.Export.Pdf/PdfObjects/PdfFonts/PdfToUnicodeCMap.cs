using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfToUnicodeCMap : PdfObject
    {
        private readonly Dictionary<int, string> CharacterMappings;
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
        public PdfToUnicodeCMap(int objectNumber, Dictionary<int, string> characterMappings, int codeSpaceMin = 0, int codeSpaceMax = 255, int bytesPerCode = 1, int version = 0) : base(objectNumber, version)
        {
            CharacterMappings = characterMappings ?? new Dictionary<int, string>();
            CodeSpaceMin = codeSpaceMin;
            CodeSpaceMax = codeSpaceMax;
            BytesPerCode = bytesPerCode;
        }

        internal override string RenderDictionary()
        {
            var cmapContent = GenerateCMapContent();
            var length = Encoding.UTF8.GetByteCount(cmapContent);

            var sb = new StringBuilder();
            sb.AppendLine(string.Format("<< /Length {0} >>", length));
            sb.AppendLine("stream");
            sb.Append(cmapContent);
            sb.Append("endstream");

            return sb.ToString();
        }

        private string GenerateCMapContent()
        {
            var sb = new StringBuilder();

            // CMap header
            sb.AppendLine("/CIDInit /ProcSet findresource begin");
            sb.AppendLine("12 dict begin");
            sb.AppendLine("begincmap");

            // CIDSystemInfo - required for ToUnicode CMaps
            sb.AppendLine("/CIDSystemInfo");
            sb.AppendLine("<< /Registry (Adobe)");
            sb.AppendLine("   /Ordering (UCS)");
            sb.AppendLine("   /Supplement 0");
            sb.AppendLine(">> def");

            // CMap name and type
            sb.AppendLine("/CMapName /Adobe-Identity-UCS def");
            sb.AppendLine("/CMapType 2 def");

            // Define codespace range
            sb.AppendLine("1 begincodespacerange");
            sb.AppendLine(string.Format("<{0}> <{1}>", FormatCode(CodeSpaceMin), FormatCode(CodeSpaceMax)));
            sb.AppendLine("endcodespacerange");

            // Generate character mappings
            if (CharacterMappings.Count > 0)
            {
                GenerateCharacterMappings(sb);
            }

            // CMap footer
            sb.AppendLine("endcmap");
            sb.AppendLine("CMapName currentdict /CMap defineresource pop");
            sb.AppendLine("end");
            sb.AppendLine("end");

            return sb.ToString();
        }

        private void GenerateCharacterMappings(StringBuilder sb)
        {
            // Group mappings into ranges where possible for efficiency
            var sortedMappings = CharacterMappings.OrderBy(kvp => kvp.Key).ToList();
            var ranges = new List<CharacterRange>();
            var individualMappings = new List<CharacterMapping>();

            // Try to identify consecutive ranges
            for (int i = 0; i < sortedMappings.Count; i++)
            {
                var current = sortedMappings[i];

                // Try to parse as simple Unicode value for range detection
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

                    // If we found a range of 3 or more, use beginbfrange
                    if (rangeEnd - rangeStart >= 2)
                    {
                        ranges.Add(new CharacterRange(rangeStart, rangeEnd, unicodeStart));
                    }
                    else
                    {
                        // Add individual mappings
                        for (int j = rangeStart; j <= rangeEnd; j++)
                        {
                            var mapping = sortedMappings.FirstOrDefault(m => m.Key == j);
                            individualMappings.Add(new CharacterMapping(j, mapping.Value));
                        }
                    }
                }
                else
                {
                    // Complex mapping (ligatures, etc.) - must be individual
                    individualMappings.Add(new CharacterMapping(current.Key, current.Value));
                }
            }

            // Output ranges
            if (ranges.Count > 0)
            {
                sb.AppendLine(string.Format("{0} beginbfrange", ranges.Count));
                foreach (var range in ranges)
                {
                    sb.AppendLine(string.Format("<{0}> <{1}> <{2}>",
                        FormatCode(range.Start),
                        FormatCode(range.End),
                        FormatUnicode(range.UnicodeStart)));
                }
                sb.AppendLine("endbfrange");
            }

            // Output individual character mappings
            if (individualMappings.Count > 0)
            {
                // Process in batches of 100 (PDF best practice)
                const int batchSize = 100;
                for (int i = 0; i < individualMappings.Count; i += batchSize)
                {
                    var batch = individualMappings.Skip(i).Take(batchSize).ToList();
                    sb.AppendLine(string.Format("{0} beginbfchar", batch.Count));
                    foreach (var mapping in batch)
                    {
                        sb.AppendLine(string.Format("<{0}> <{1}>",
                            FormatCode(mapping.Code),
                            mapping.Unicode));
                    }
                    sb.AppendLine("endbfchar");
                }
            }
        }

        private bool TryParseSimpleUnicode(string hexString, out int unicodeValue)
        {
            unicodeValue = 0;

            // Check if it's a simple 4-digit hex Unicode value (e.g., "0041" for 'A')
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
                // Default to 2 bytes
                return code.ToString("X4");
            }
        }

        private string FormatUnicode(int unicodeValue)
        {
            // Format Unicode value in UTF-16BE hex format
            // For BMP characters (U+0000 to U+FFFF), use 4 hex digits
            if (unicodeValue <= 0xFFFF)
            {
                return unicodeValue.ToString("X4");
            }
            else
            {
                // For supplementary characters, convert to UTF-16 surrogate pair
                unicodeValue -= 0x10000;
                int highSurrogate = 0xD800 + (unicodeValue >> 10);
                int lowSurrogate = 0xDC00 + (unicodeValue & 0x3FF);
                return string.Format("{0:X4}{1:X4}", highSurrogate, lowSurrogate);
            }
        }

        // Helper classes to replace tuples
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
