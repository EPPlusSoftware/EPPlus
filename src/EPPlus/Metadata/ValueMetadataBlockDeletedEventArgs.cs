using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class ValueMetadataBlockDeletedEventArgs : EventArgs
    {
        public ValueMetadataBlockDeletedEventArgs(uint valueMetadataBlockId)
        {
            ValueMetadataBlockId = valueMetadataBlockId;
        }

        public uint ValueMetadataBlockId
        {
            get;
        }
    }
}
