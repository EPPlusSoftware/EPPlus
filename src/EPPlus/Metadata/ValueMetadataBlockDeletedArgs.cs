using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class ValueMetadataBlockDeletedArgs : EventArgs
    {
        public uint ValueMetadataBlockId
        {
            get; set;
        }
    }
}
