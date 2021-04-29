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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Core.ExternalReferences
{
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
                if (reader.LocalName == "fallback")
                {
                    reader.ReadElementContentAsString();
                    continue;
                }
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "ddeItem")
                {
                    DdeItems.Add(new ExcelExternalDdeItem()
                    {
                        Name = reader.GetAttribute("name"),
                        Advise = XmlHelper.GetBoolFromString(reader.GetAttribute("advice")),
                        Icon = XmlHelper.GetBoolFromString(reader.GetAttribute("icon")),
                        PreferPicture = XmlHelper.GetBoolFromString(reader.GetAttribute("preferPic")),
                    });
                    break;
                }
            }
        }

        public override eExternalLinkType ExternalLinkType
        {
            get
            {
                return eExternalLinkType.DdeLink;
            }
        }
        public string DdeService { get; set; }
        public string DdeTopic { get; set; }

        public ExcelExternalDdeItemCollection DdeItems
        {
            get;
        } = new ExcelExternalDdeItemCollection();
        internal override void Save(StreamWriter sw)
        {

        }
    }
}

