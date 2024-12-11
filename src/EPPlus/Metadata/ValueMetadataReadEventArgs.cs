using System;

namespace OfficeOpenXml.Metadata
{
    internal class ValueMetadataReadEventArgs : EventArgs
    {
        public ValueMetadataReadEventArgs(uint id, uint oneBasedIndex)
        {
            Id = id;
            OneBasedIndex = oneBasedIndex;
        }

        public uint Id { get; }

        public uint OneBasedIndex {  get; }
    }
}
