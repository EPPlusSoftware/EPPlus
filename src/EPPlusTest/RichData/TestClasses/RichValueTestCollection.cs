using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.RichData.TestClasses
{
    internal class RichValueTestCollection : IndexedCollection<RichValueTest>
    {
        public RichValueTestCollection(RichDataIndexStore store) : base(store, RichDataEntities.RichValue)
        {
        }

        public override RichDataEntities EntityType => RichDataEntities.RichValue;
    }
}
