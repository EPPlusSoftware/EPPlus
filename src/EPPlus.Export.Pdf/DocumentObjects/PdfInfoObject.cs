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
using EPPlus.Export.Pdf.Settings;
using System;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfInfoObject : PdfObject
    {
        internal string Title;
        internal string LicenseType;
        internal string LicenseHolder;
        internal string Author;

        public PdfInfoObject(int objectNumber, PdfDocumentSettings documentSettings, int version = 0) : base(objectNumber, version)
        {
            Title = documentSettings.Title;
            Author = documentSettings.Author;
            LicenseType = documentSettings.LicenseType;
            LicenseHolder = documentSettings.LicenseHolder;
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
                            $"   /Author ({Author})\n" +
                            $"   /Subject (EPPlus PDF Export with {LicenseType} for {LicenseHolder})\n" +
                            $"   /Keywords (EPPlus, EPPlus Software, PDF, Export, {LicenseType}, {LicenseHolder})" +
                            $"   /Creator (EPPlus Software {LicenseType}, {LicenseHolder})\n" +
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
                           $"   /Author ({Author})\n" +
                           $"   /Subject (EPPlus PDF Export with {LicenseType} for {LicenseHolder})\n" +
                           $"   /Keywords (EPPlus, EPPlus Software, PDF, Export, {LicenseType}, {LicenseHolder})" +
                           $"   /Creator (EPPlus Software {LicenseType}, {LicenseHolder})\n" +
                           $"   /Producer (EPPlus Software PDF Exporter)\n" +
                           $"   /CreationDate ({pdfDate})\n" +
                           $"   /ModDate ({pdfDate})\n" +
                           $"   /Trapped /False >>");
        }
    }
}