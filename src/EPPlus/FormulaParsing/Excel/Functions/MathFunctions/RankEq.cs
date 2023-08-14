using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the Mode (the most frequently occurring value) of a list of supplied numbers (if more than one value has same rank, the top rank of that set is returned)")]
    internal class RankEq : Rank
    {
        public override string NamespacePrefix => "_xlfn.";
    }
}
