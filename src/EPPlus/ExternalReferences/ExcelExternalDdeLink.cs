/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents an external DDE link.
    /// </summary>
    public class ExcelExternalDdeLink : ExcelExternalLink
    {
        internal ExcelExternalDdeLink(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement) : base (wb, reader, part, workbookElement)
        {
            DdeService = reader.GetAttribute("ddeService");
            DdeTopic = reader.GetAttribute("ddeTopic");
            
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "ddeItems":
                            ReadDdeItems(reader);
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name == "ddeLink")
                    {
                        break;
                    }
                }
            }
        }
        private void ReadDdeItems(XmlTextReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Fallback")
                {
                    XmlStreamHelper.ReadUntil(reader, "Fallback");
                }
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "ddeItem")
                {
                    DdeItems.Add(new ExcelExternalDdeItem()
                    {
                        Name = reader.GetAttribute("name"),
                        Advise = XmlHelper.GetBoolFromString(reader.GetAttribute("advise")),
                        Ole = XmlHelper.GetBoolFromString(reader.GetAttribute("ole")),
                        PreferPicture = XmlHelper.GetBoolFromString(reader.GetAttribute("preferPic")),
                    });
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName=="ddeItems")
                {
                    break;
                }
            }
        }

        /// <summary>
        /// The type of external link
        /// </summary>
        public override eExternalLinkType ExternalLinkType
        {
            get
            {
                return eExternalLinkType.DdeLink;
            }
        }
        /// <summary>
        /// Service name for the DDE connection
        /// </summary>
        public string DdeService { get; internal set; }
        /// <summary>
        /// Topic for DDE server. 
        /// </summary>
        public string DdeTopic { get; internal set; }

        /// <summary>
        /// A collection of <see cref="ExcelExternalDdeItem" />
        /// </summary>
        public ExcelExternalDdeItemCollection DdeItems
        {
            get;
        } = new ExcelExternalDdeItemCollection();
        internal override void Save(StreamWriter sw)
        {
            sw.Write($"<ddeLink ddeTopic=\"{DdeTopic}\" ddeService=\"{DdeService}\"><ddeItems>");
            foreach (ExcelExternalDdeItem item in DdeItems)
            {                
                sw.Write(string.Format("<ddeItem name=\"{0}\" {1}{2}{3}/>",
                  item.Name,
                  item.Advise.GetXmlAttributeValue("advise", false),
                  item.Ole.GetXmlAttributeValue("ole", false),
                  item.PreferPicture.GetXmlAttributeValue("preferPic", false)));
            }
            sw.Write("</ddeItems></ddeLink>");
        }
    }
}

