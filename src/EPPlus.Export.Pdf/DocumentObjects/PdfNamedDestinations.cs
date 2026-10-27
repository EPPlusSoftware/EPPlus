/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfNamedDestinations : PdfObject
    {
        internal List<(string Name, int PageObj, double X, double Y)> Dests = new List<(string, int, double, double)>();

        public PdfNamedDestinations(int objectNumber, int version = 0)
            : base(objectNumber, version)
        { }

        private string Body()
        {
            var sb = new StringBuilder();
            sb.Append("<< /Dests << /Names [ ");
            foreach (var d in Dests.OrderBy(d => d.Name, System.StringComparer.Ordinal))
            {
                sb.Append(d.Name.ToPdfLiteral());
                sb.Append(" [ ");
                sb.Append(d.PageObj.ToPdfStringF0());
                sb.Append(" 0 R /XYZ ");
                sb.Append(d.X.ToPdfStringF4());
                sb.Append(' ');
                sb.Append(d.Y.ToPdfStringF4());
                sb.Append(" null ] ");
            }
            sb.Append("] >> >>");
            return sb.ToString();
        }

        internal override string RenderDictionary() => Body();
        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Body());
    }
}