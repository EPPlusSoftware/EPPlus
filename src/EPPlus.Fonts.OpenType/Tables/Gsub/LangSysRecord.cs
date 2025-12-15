using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LangSysRecord
    {
        // USHORT LangSysTag (TAG, 4 bytes, men lagras som uint)
        public uint LangSysTag { get; set; }

        // Referens till det faktiska LangSysTable objektet
        public LangSysTable LangSysTable { get; set; }
    }
}
