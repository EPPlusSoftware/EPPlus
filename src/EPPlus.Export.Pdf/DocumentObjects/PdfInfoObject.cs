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
using System;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
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
            DateTime now = DateTime.Now;
            TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(now);
            string sign = offset < TimeSpan.Zero ? "-" : "+";
            offset = offset.Duration();
            string pdfDate = string.Format("D:{0:yyyyMMddHHmmss}{1}{2:00}'{3:00}'", now, sign, offset.Hours, offset.Minutes);
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Title ({Title})\n" +
                            $"   /Author (EPPlus)\n" +
                            $"   /Subject (EPPlus PDF Export)\n" +
                            $"   /Keywords (EPPlus, EPPlus Software)" +
                            $"   /Creator (EPPlus Software)\n" +
                            $"   /Producer (EPPlus Software PDF Exporter)\n" +
                            $"   /CreationDate ({pdfDate})\n" +
                            $"   /ModDate ({pdfDate})\n" +
                            $"   /Trapped /False >>");
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            DateTime now = DateTime.Now;
            TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(now);
            string sign = offset < TimeSpan.Zero ? "-" : "+";
            offset = offset.Duration();
            string pdfDate = string.Format("D:{0:yyyyMMddHHmmss}{1}{2:00}'{3:00}'", now, sign, offset.Hours, offset.Minutes);
            WriteAscii(bw, $"<< /Title ({Title})\n" +
                           $"   /Author (EPPlus)\n" +
                           $"   /Subject (EPPlus PDF Export)\n" +
                           $"   /Keywords (EPPlus, EPPlus Software)\n" +
                           $"   /Creator (EPPlus Software)\n" +
                           $"   /Producer (EPPlus Software PDF Exporter)\n" +
                           $"   /CreationDate ({pdfDate})\n" +
                           $"   /ModDate ({pdfDate})\n" +
                           $"   /Trapped /False >>");
        }
    }
}
