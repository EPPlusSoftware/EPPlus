/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Patterns
{
    internal  class PdfShadingPattern : PdfPattern
    {
        internal int shadingObjectNumber;
        internal double[] Matrix;
        //internal ExtGState dictionary //implement later

        public PdfShadingPattern(int objectNumber, int shadingObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.shadingObjectNumber = shadingObjectNumber;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Pattern\n" +
                            $"   /PatternType 2\n" +
                            $"   /Shading {shadingObjectNumber.ToPdfStringF0()} 0 R");
            if (Matrix != null)
            {
                var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfStringF4()).ToArray());
                sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            }
            sb.Append(" >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Pattern\n" +
                            $"   /PatternType 2\n" +
                            $"   /Shading {shadingObjectNumber.ToPdfStringF0()} 0 R");
            if (Matrix != null)
            {
                var matrixStr = string.Join(" ", Matrix.Select(w => w.ToPdfStringF4()).ToArray());
                sb.AppendFormat($"\n   /Matrix [ {matrixStr} ]");
            }
            sb.Append(" >>");
            WriteAscii(bw, sb.ToString());
        }
    }
}
