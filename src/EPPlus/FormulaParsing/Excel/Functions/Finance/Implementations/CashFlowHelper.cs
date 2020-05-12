using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class CashFlowHelper
    {
        private static double GetAnnuityFactor(double rate, double nper, PmtDue type)
        {
            return rate == 0.0 ? nper : (1 + rate * (int)type) * (1 - 1d / System.Math.Pow(1.0 + rate, nper)) / rate;
        }

        private static double FvCalc(double rate, double nper, double pmt, double pv, PmtDue type)
        {
            return -1 * (pv * System.Math.Pow(1.0 + rate, nper) + pmt * (GetAnnuityFactor(rate, nper, type) * System.Math.Pow(1.0 + rate, nper)));
        }

        /// <summary>
        /// The Excel FV function calculates the Future Value of an investment with periodic constant payments and a constant interest rate.
        /// </summary>
        /// <param name="rate">The interest rate, per period.</param>
        /// <param name="nper">The number of periods for the lifetime of the annuity.</param>
        /// <param name="pmt">An optional argument that specifies the payment per period.</param>
        /// <param name="pv">An optional argument that specifies the present value of the annuity - i.e. the amount that a series of future payments is worth now.</param>
        /// <param name="type">An optional argument that defines whether the payment is made at the start or the end of the period.</param>
        /// <returns></returns>
        public static double Fv(double rate, double nper, double pmt = 0d, double pv = 0d, PmtDue type = 0)
        {
            if((type == PmtDue.EndOfPeriod && rate == -1d))
                return -(pv * System.Math.Pow(1d + rate, nper));
            if (rate == -1d && type == PmtDue.EndOfPeriod) return -(pv * System.Math.Pow(1d + rate, nper) + pmt);
            return FvCalc(rate, nper, pmt, pv, type); ;
        }

        /// <summary>
        /// Calculates the present value
        /// </summary>
        /// <param name="rate">The interest rate, per period.</param>
        /// <param name="nper">The number of periods for the lifetime of the annuity or investment.</param>
        /// <param name="pmt">An optional argument that specifies the payment per period.</param>
        /// <param name="fv">An optional argument that specifies the future value of the annuity, at the end of nper payments.If the[fv] argument is omitted, it takes on the default value 0.</param>
        /// <param name="type">An optional argument that defines whether the payment is made at the start or the end of the period. See <see cref="PmtDue"></see></param>
        /// <returns></returns>
        public static double Pv(double rate, double nper, double pmt = 0d, double fv = 0d, PmtDue type = 0)
        {
            return -1 * (fv * (1d / System.Math.Pow(1.0 + rate, nper)) + pmt * GetAnnuityFactor(rate, nper, type));
        }

        /// <summary>
        /// The Excel NPV function calculates the Net Present Value of an investment, based on a supplied discount rate, and a series of future payments and income.
        /// </summary>
        /// <param name="rate">The discount rate over one period.</param>
        /// <param name="payments">Numeric values, representing a series of regular payments and income</param>
        /// <returns></returns>
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
