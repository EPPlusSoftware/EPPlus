using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal static class IdGenerator
    {
        private static int _nextId = 0;
        public static int GetNewId()
        {
            return ++_nextId;
        }
    }
}
