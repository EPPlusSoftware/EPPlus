using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexedValue : IdentityItem
    {
        public IndexedValue(int originalValue, RichDataEntities entity)
        {
            OriginalValue = originalValue;
            EntityType = entity;
        }

        public RichDataEntities EntityType { get; private set; }

        public RichDataEntities DestinationEntity { get; private set; }
        public int OriginalValue { get; private set; }
        public int CurrentValue { get; set; }
    }
}
