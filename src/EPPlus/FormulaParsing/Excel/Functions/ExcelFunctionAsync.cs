using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils.RemoteCalls;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal abstract class ExcelFunctionAsync : ExcelFunction
    {
        public abstract CompileResult Complete(RemoteTask task);
    }
}
