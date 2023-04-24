using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.XMLWritingEncoder
{
    internal static class AttributeEncoder
    {
        /// <summary>
        /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string EncodeXMLAttribute(this string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }

    }
}
