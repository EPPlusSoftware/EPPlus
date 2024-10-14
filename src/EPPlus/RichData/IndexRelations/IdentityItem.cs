using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal abstract class IdentityItem
    {
        protected IdentityItem()
        {
            _id = IdGenerator.GetNewId();
        }

        private readonly int _id;
        public int Id => _id;
    }
}
