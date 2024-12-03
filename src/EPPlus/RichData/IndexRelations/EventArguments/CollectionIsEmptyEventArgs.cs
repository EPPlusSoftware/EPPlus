using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations.EventArguments
{
    internal class CollectionIsEmptyEventArgs : EventArgs
    {
        public CollectionIsEmptyEventArgs(RelationDeletions deletions)
        {
            Deletions = deletions;
        }

        public RelationDeletions Deletions { get; }
    }
}
