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

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal abstract class PdfObject
    {
        internal int objectNumber;
        internal int version;

        public PdfObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
            this.version = version;
        }

        internal string RenderBody()
        {
            var w = new PdfWriter(null, captureDebug: true);
            WriteBody(w);
            return w.DebugText;
        }

        public void Write(PdfWriter w)
        {
            w.Text($"{objectNumber.ToPdfStringF0()} {version.ToPdfStringF0()} obj\n");
            WriteBody(w);
            w.Text("\nendobj\n");
        }

        protected abstract void WriteBody(PdfWriter w);
    }
}
