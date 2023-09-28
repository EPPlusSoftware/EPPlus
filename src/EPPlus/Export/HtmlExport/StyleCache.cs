using System;
using System.Collections.Generic;
using System.Diagnostics.SymbolStore;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class StyleCache : Dictionary<string, int>
    {
        internal StyleCache()
        {
            //cache = new Dictionary<string, int>();
        }

        internal bool IsAdded(string key, out int id)
        {
            if (ContainsKey(key))
            {
                id = base[key];
                return true;
            }
            else
            {
                id = Count + 1;
                Add(key, id);
                return false;
            }
        }


    }
}
