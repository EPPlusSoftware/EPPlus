using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class PmtProvider : IPmtProvider
    {
        public double GetPmt(double Rate, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            return InternalMethods.PMT_Internal(Rate, NPer, PV, FV, Due).Result;
        }
    }
}
