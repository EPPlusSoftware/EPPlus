using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class ExtensionSubstSubTable : FontTableElement
    {
        public ushort ExtensionLookupType { get; set; }
        public FontTableElement InnerSubTable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }

        // Vi behöver inte skriva ner denna som en Extension i subset-fonten sen,
        // vi kan "packa upp" den och spara den som en vanlig LigatureSubst (Typ 4).
    }
}
