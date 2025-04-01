using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.Utils.RemoteCalls
{
    internal abstract class RemoteTask
    {
        protected RemoteTask(ExcelFunctionAsync func, ParsingContext ctx)
        {
            Id = Guid.NewGuid();
            _func = func;
            _ctx = ctx;
        }

        private readonly ExcelFunctionAsync _func;

        private readonly ParsingContext _ctx;

        protected ExcelFunctionAsync ExcelFunction => _func;

        internal ParsingContext ParsingContext => _ctx;

        public Guid Id { get; }

        public bool Done { get; set; }

        public abstract void DoWork();
    }
}
