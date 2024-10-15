using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexRelation : IdentityItem
    {
        public IndexRelation(IndexEndpoint from, IndexEndpoint to, IndexType indexType)
        {
            From = from;
            To = to;
            IndexType = indexType;
        }
        public IndexEndpoint From { get; set; }

        public IndexEndpoint To { get; set; }

        public IndexType IndexType { get; set; }
    }
}
