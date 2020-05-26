using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class YearFracProvider : IYearFracProvider
    {
        public YearFracProvider(ParsingContext context)
        {
            _context = context;
        }

        private readonly ParsingContext _context;
        public double GetYearFrac(System.DateTime date1, System.DateTime date2, DayCountBasis basis)
        {
            var func = new Yearfrac();
            var args = new List<FunctionArgument> { new FunctionArgument(date1.ToOADate()), new FunctionArgument(date2.ToOADate()), new FunctionArgument((int)basis) };
            var result = func.Execute(args, _context);
            return result.ResultNumeric;
        }
    }
}
