using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class FvProvider : IFvProvider
    {
        public double GetFv(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            return InternalMethods.FV_Internal(Rate, NPer, Pmt, PV, Due);
        }
    }
}
