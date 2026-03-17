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
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects
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

        public virtual string ToPdfString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0} {1} obj\n", objectNumber, version);
            sb.Append(RenderDictionary());
            sb.Append("\nendobj\n");
            return sb.ToString();
        }

        public virtual byte[] ToPdfBytes()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0} {1} obj\n", objectNumber, version);
            sb.Append(RenderDictionary());
            sb.Append("\nendobj\n");
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        public virtual void ToPdfBytes(BinaryWriter bw)
        {
            var sb = new StringBuilder();
            WriteAscii(bw, $"{objectNumber} {version} obj\n");
            RenderDictionary(bw);
            WriteAscii(bw, $"\nendobj\n");
        }

        protected static void WriteAscii(BinaryWriter bw, string s)
        {
            bw.Write(Encoding.ASCII.GetBytes(s));
        }

        internal abstract string RenderDictionary();

        internal abstract void RenderDictionary(BinaryWriter bw);
    }
}
