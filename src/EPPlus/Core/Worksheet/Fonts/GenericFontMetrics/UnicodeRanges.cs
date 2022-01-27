using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal class UniCodeRange
    {
        public UniCodeRange(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; set; }

        public int End { get; set; }

        public bool IsInRange(int c)
        {
            return (c >= Start && c <= End);
        }

        public IEnumerable<char> ToCharList()
        {
            var result = new List<char>();
            for (var c = Start; c <= End; c++)
            {
                result.Add(Convert.ToChar(c));
            }
            return result;
        }

        /// <summary>
        /// Unicode ranges to cover Japanese/Kanji characters
        /// </summary>
        public static IEnumerable<UniCodeRange> JapaneseKanji
        {
            get
            {
                return new List<UniCodeRange>
                {
                    // Hiragana
                    new UniCodeRange(0x3040, 0x3096),
                    // Katakana
                    new UniCodeRange(0x30A0, 0x30FF),
                    // Kanji
                    new UniCodeRange(0x3400, 0x4DB5),
                    new UniCodeRange(0x4E00, 0x9FCB),
                    new UniCodeRange(0xF900, 0xFA6A),
                    // Kanji Radicals
                    new UniCodeRange(0x2E80, 0x2FD5),
                    // Katakana and Punctuation (Half Width)
                    new UniCodeRange(0xFF5F, 0xFF9F),
                    // Japanese Symbols and Punctuation
                    new UniCodeRange(0x3000, 0x303F),
                    // Miscellaneous Japanese Symbols and Characters
                    new UniCodeRange(0x31F0, 0x31FF),
                    new UniCodeRange(0x3220, 0x3243),
                    new UniCodeRange(0x3280, 0x337F),
                    // Alphanumeric and Punctuation (Full Width)
                    new UniCodeRange(0xFF01, 0xFF5E)
                };
            }
        }
    }
}
