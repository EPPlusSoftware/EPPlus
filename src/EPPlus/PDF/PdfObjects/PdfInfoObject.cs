using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfInfoObject : PdfObject
    {
        public string Title;

        public PdfInfoObject(int objectNumber, string Title, int version = 0) : base(objectNumber, version)
        {
            this.Title = Title;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Author (EPPlus)\n" +
                            $"   /CreationDate ({DateTime.Now.ToString()})\n" +
                            $"   /ModDate ({DateTime.Now.ToString()})\n" +
                            $"   /Producer (EPPlus PDF Exporter)\n" +
                            $"   /Title ({Title}) >>");
            return sb.ToString();
        }
    }
}
