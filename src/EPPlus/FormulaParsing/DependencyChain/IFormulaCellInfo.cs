using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public interface IFormulaCellInfo
    {
        string Worksheet { get; }

        string Address { get; }

        string Formula { get; }
    }
}
