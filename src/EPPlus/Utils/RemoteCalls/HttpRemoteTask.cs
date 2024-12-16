using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.RemoteCalls
{
    internal class HttpRemoteTask : RemoteTask
    {
        public HttpRemoteTask(string url, ExcelFunctionAsync func, ParsingContext ctx)
            : base(func, ctx)
        {
            Url = url;
        }

        public string Url { get; private set; }

        public byte[] ResponseBytes { get; set; }

        public override void DoWork()
        {
            // do the work here
            var cr = ExcelFunction.Complete(this);
            // todo: put the compile result on a queue that can be picked up by the calc engine.
            throw new NotImplementedException();
        }
    }
}
