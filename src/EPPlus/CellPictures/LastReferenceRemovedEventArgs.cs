using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
