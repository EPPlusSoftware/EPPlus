using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.Interfaces
{
    /// <summary>
    /// Memory settings interface
    /// </summary>
    public interface IMemorySettings
    {
        IMemoryManager MemoryManager { get; set;  }

        bool UseRecyclableMemory { get; set; }
    }
}
