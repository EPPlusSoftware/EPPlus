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
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Helpers;
using System;

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

        private string Body()
        {
            DateTime now = DateTime.Now;
            TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(now);
            string sign = offset < TimeSpan.Zero ? "-" : "+";
            offset = offset.Duration();
            string pdfDate = string.Format("D:{0:yyyyMMddHHmmss}{1}{2:00}'{3:00}'", now, sign, offset.Hours, offset.Minutes);
            string subject = "EPPlus PDF Export with " + LicenseType + " for " + LicenseHolder;
            string keywords = "EPPlus, EPPlus Software, PDF, Export, " + LicenseType + ", " + LicenseHolder;
            string creator = "EPPlus Software " + LicenseType + ", " + LicenseHolder;

            return $"<< /Title {Title.ToPdfTextString()}\n" +
                   $"   /Author {Author.ToPdfTextString()}\n" +
                   $"   /Subject {subject.ToPdfTextString()}\n" +
                   $"   /Keywords {keywords.ToPdfTextString()}\n" +
                   $"   /Creator {creator.ToPdfTextString()}\n" +
                   $"   /Producer (EPPlus Software PDF Exporter)\n" +
                   $"   /CreationDate ({pdfDate})\n" +
                   $"   /ModDate ({pdfDate})\n" +
                   $"   /Trapped /False >>";
        }

        protected override void WriteBody(PdfWriter w) => w.Text(Body());
    }
}