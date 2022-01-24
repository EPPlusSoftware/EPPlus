using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.FontLocalization
{
    public class LanguageMapping
    {
        public int code { get; set; }

        public Languages Language { get; set; }

        public static LanguageMapping Create(int code, Languages language)
        {
            return new LanguageMapping
            {
                code = code,
                Language = language
            };
        }

        public override string ToString()
        {
            return Language.ToString();
        }
    }
}
