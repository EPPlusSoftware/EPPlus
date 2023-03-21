using OfficeOpenXml.Core.CellStore;
using System;

namespace OfficeOpenXml.FormulaParsing
{
    internal class DynamicArrayFormulaOutput
    {
        internal static void FillDynamicArrayFromRangeInfo(RpnFormula f, IRangeInfo ri, RangeHashset rd, RpnOptimizedDependencyChain depChain)
        {
            f._isDynamic = true;
        }
    }
}