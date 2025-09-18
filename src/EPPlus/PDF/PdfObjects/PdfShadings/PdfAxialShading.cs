using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects.PdfFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfShadings
{
    internal class PdfAxialShading : PdfShading
    {
        internal double[] Coords;
        internal double[] Domain = null;
        internal PdfFunction Function;
        internal bool[] Extend = null;

        public PdfAxialShading(int objectNumber, int version = 0) : base(objectNumber, version) { }

        internal override string RenderDictionary()
        {
            var coordsStr = string.Join(" ", Coords.Select(w => w.ToPdfString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /ShadingType 2\n" +
                            $"   /ColorSpace {ColorSpace.ToString()}\n" +
                            $"   /Coords [{coordsStr}]\n");
            sb.AppendFormat($"   /Function {Function.RenderDictionary()}");
            if (Domain != null)
            {
                var domainStr = string.Join(" ", Domain.Select(w => w.ToPdfString()).ToArray());
                sb.AppendFormat($"\n   /Domain [{domainStr}]");
            }
            if(Extend != null)
            {
                var extendStr = string.Join(" ", Extend.Select(w => w.ToString()).ToArray());
                sb.AppendFormat($"\n   /Extend [{extendStr}]");
            }
            sb.Append(">>");
            return sb.ToString();
        }
    }
}
