using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal interface IndexedCollectionInterface
    {
        RichDataEntities EntityType { get; }

        int Count { get; }

        IndexEndpoint this[int index] { get; }

        void DeleteEndpoint(int id);

        int GetNextIndex();
    }
}
