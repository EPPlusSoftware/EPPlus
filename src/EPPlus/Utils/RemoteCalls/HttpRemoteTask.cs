using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Services;
using OfficeOpenXml.RichData.RichValues.WebImages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.RemoteCalls
{
    internal class HttpRemoteTask : RemoteTask
    {
        public HttpRemoteTask(string url, ExcelFunctionAsync func, ParsingContext ctx, WebImageSizing sizing)
            : base(func, ctx)
        {
            Url = url;
            Cell = ctx.CurrentCell;
            Sizing = sizing;
        }
        public string Url { get; private set; }

        public byte[] ResponseBytes { get; set; }
        public WebImageSizing Sizing { get; private set; }
        public FormulaCellAddress Cell { get; private set; }
        public override void DoWork()
        {
            try
            {
                ResponseBytes = ParsingContext.Package.Settings.ImageFunctionService.Download(Url);
                // do the work here            
                var cr = ExcelFunction.Complete(this);
                ParsingContext.Package.Workbook.Worksheets[Cell.WorksheetIx].SetValueInner(Cell.Row, Cell.Column, cr.ResultValue);

            }
            catch(Exception ex)
            {
                ParsingContext.Package.Workbook.Worksheets[Cell.WorksheetIx].SetValueInner(Cell.Row, Cell.Column, ErrorValues.ValueError);
            }
            finally
            {
                ParsingContext.RemoteCallManager.TaskComplate(this);
            }
        }
    }
}
