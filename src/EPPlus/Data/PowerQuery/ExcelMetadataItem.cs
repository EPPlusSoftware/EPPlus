/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// A power query metadata item used for describing the M-Formula.
    /// </summary>
    public class ExcelPowerQueryMetadataItem
    {
        internal ExcelPowerQueryMetadataItem(XmlNamespaceManager nsm, XmlNode topNode, CultureInfo culture)
        {
            var xh = XmlHelperFactory.Create(nsm, topNode);
            ItemType = xh.GetXmlEnum("ItemLocation/ItemType", ePowerQueryMetadataItemType.Formula);
            ItemPath = xh.GetXmlNodeString("ItemLocation/ItemPath");

            foreach (XmlElement node in xh.GetNodes("StableEntries/Entry"))
            {
                var type = node.GetAttribute("Type");
                var sv = node.GetAttribute("Value");
                object value;
                var prefix = sv[0];
                var s = sv.Substring(1);
                switch (prefix)
                {
                    case 's':
                    case 'S':
                        value = s;
                        break;
                    case 'l':
                    case 'L':
                        if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long l))
                        {
                            value = l;
                        }
                        else
                        {
                            value = s;
                        }
                        break;
                    case 'b':
                    case 'B':
                        if (bool.TryParse(s, out bool b))
                        {
                            value = b;
                        }
                        else
                        {
                            value = s;
                        }
                        break;
                    case 'd':
                    case 'D':
                        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTime dt))
                        {
                            value = dt;
                        }
                        else
                        {
                            value = s;
                        }
                        break;
                    case 'f':
                    case 'F':
                        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double f))
                        {
                            value = f;
                        }
                        else
                        {
                            value = s;
                        }
                        break;
                    case 'c':
                    case 'C':
#if (NET8_0_OR_GREATER)
                        if (Guid.TryParse(s, CultureInfo.InvariantCulture, out Guid guid))
                        {
                            value = guid;
                        }
                        else
                        {
                            value = s;
                        }
#else                        
                        try
                        {
                            value = new Guid(s);   
                        }
                        catch
                        {
                            value = s;
                        }
#endif
                        break;
                    default:
                        value = sv;
                        break;
                }
                Entries.Add(new ExcelPowerQueryMetaDataEntry(type, value, true, false));
            }
        }
        /// <summary>
        /// If the items applies to all formulas or only a single formula. The formula must exist in the <see cref="ExcelPowerQuerySettings.Formulas"/>.
        /// </summary>
        public ePowerQueryMetadataItemType ItemType { get; set; }
        /// <summary>
        /// The item path if <see cref="ItemType"/> is set to <see cref="ePowerQueryMetadataItemType.Formula"/>.
        /// </summary>
        public string ItemPath { get; set; }
        /// <summary>
        /// A collection of metadata entries.
        /// </summary>
        public List<ExcelPowerQueryMetaDataEntry> Entries
        {
            get;
        } = new List<ExcelPowerQueryMetaDataEntry>();
    }
}