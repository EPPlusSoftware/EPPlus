using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    public class SignatureTime
    {
        string Format;
        DateTime Value;
        string Id;

        string packagePrefix = "mdssi";
        string packageNameSpace = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        string strValue;

        internal SignatureTime(DateTime value, string format = "YYYY-MM-DDThh:mm:ssTZD") 
        {
            Format = format;
            Value = value;
            string toStringFormat = format == "YYYY-MM-DDThh:mm:ssTZD" ? "yyyy-MM-ddTHH:mm:ssZ" : format;

            strValue = value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<{packagePrefix}:SignatureTime xmlns:{packagePrefix}=\"{packageNameSpace}\">");
            sb.Append($"<{packagePrefix}:Format>{Format}</{packagePrefix}:Format>");
            sb.Append($"<{packagePrefix}:Value>{strValue}</{packagePrefix}:Value>");
            sb.Append($"</{packagePrefix}:SignatureTime>");

            return sb.ToString();
        }
    }
}
