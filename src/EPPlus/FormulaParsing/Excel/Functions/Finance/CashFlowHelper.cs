using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    internal static class CashFlowHelper
    {

        private static double FvCalc(double rate, double nper, double pmt, double pv, int pd)
        {
            var pvFactor = rate == 0.0 ? nper : (1 + rate * pd) * (1 - 1d / System.Math.Pow(1.0 + rate, nper)) / rate;
            return -(pv * System.Math.Pow(1.0 + rate, nper) + pmt * (pvFactor * System.Math.Pow(1.0 + rate, nper)));
        }

        public static double Fv(double rate, double nper, double pmt = 0d, double pv = 0d, int type = 0)
        {
            if((type == 1 && rate == -1d))
                return -(pv * System.Math.Pow(1d + rate, nper));
            return (rate != -1.0 ? 0 : (type == 0 ? 1 : 0)) != 0 ? -(pv * System.Math.Pow(1.0 + rate, nper) + pmt) : FvCalc(rate, nper, pmt, pv, type);
        }
        public static double Npv(double rate, IEnumerable<double> payments)
        {
            var retVal = 0d;
            for (var x = 0; x < payments.Count(); x++)
            {
                var payment = payments.ElementAt(x);
                retVal += payment * (1d / System.Math.Pow(1d + rate, x + 1));
            }
            return retVal;
        }
    }
}
