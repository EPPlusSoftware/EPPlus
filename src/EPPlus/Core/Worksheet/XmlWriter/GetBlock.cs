/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2023         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Core.Worksheet.XmlWriter
{
    internal static class GetBlock
    {
        internal static void Pos(string xml, string tag, ref int start, ref int end)
        {
            Match startmMatch, endMatch;
            startmMatch = Regex.Match(xml.Substring(start), string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

            if (!startmMatch.Success) //Not found
            {
                start = -1;
                end = -1;
                return;
            }
            var startPos = startmMatch.Index + start;
            if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
            {
                end = startPos + startmMatch.Length;
            }
            else
            {
                endMatch = Regex.Match(xml.Substring(start), string.Format("(</[^>]*{0}[^>]*>)", tag));
                if (endMatch.Success)
                {
                    end = endMatch.Index + endMatch.Length + start;
                }
            }
            start = startPos;
        }
    }
}
