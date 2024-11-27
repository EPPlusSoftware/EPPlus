using System;

namespace OfficeOpenXml.CellPictures
{
    internal class LastReferenceRemovedEventArgs : EventArgs
    {
        public LastReferenceRemovedEventArgs(uint vmId)
        {
            VmId = vmId;
        }

        public uint VmId { get; }
    }
}
