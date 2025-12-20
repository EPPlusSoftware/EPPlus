using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    internal class LookupRewriteResult
    {
        public LookupListTable NewLookupList { get; set; }
        public Dictionary<int, int> OldToNewIndexMap { get; set; }
    }
}
