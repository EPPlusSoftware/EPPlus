/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/30/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Utils
{
    internal class XmlStreamHelper
    {
        internal static void ReadUntil(XmlTextReader reader, string elementName)
        {
            do { if (!reader.Read()) break; } while (!(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == elementName));
        }
    }
}
