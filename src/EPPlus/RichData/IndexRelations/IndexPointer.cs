using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal abstract class IndexPointer : IdentityItem
    {
        public IndexPointer(int originalValue, RichDataEntities entity, RichDataEntities destinationEntity)
        {
            OriginalValue = originalValue;
            EntityType = entity;
            DestinationEntity = destinationEntity;
        }

        public RichDataEntities EntityType { get; private set; }

        public RichDataEntities DestinationEntity { get; private set; }
        public int OriginalValue { get; private set; }
        public int CurrentValue { get; set; }
    }
}
