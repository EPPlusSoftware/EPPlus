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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents an external DDE link.
    /// </summary>
    public class ExcelExternalOleLink : ExcelExternalLink
    {
        internal ExcelExternalOleLink(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement) : base(wb, reader, part, workbookElement)
        {
            ExternalOleXml = new XmlDocument();
            ExternalOleXml.Load(part.GetStream());
            var rId = reader.GetAttribute("id", ExcelPackage.schemaRelationships);
            if(!string.IsNullOrEmpty(rId))
            {
                Relation = part.GetRelationship(rId);
            }
            ProgId = reader.GetAttribute("progId");
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "oleItems":
                            ReadOleItems(reader);
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Name == "oleLink")
                    {
                        break;
                    }
                }
            }
        }
        private void ReadOleItems(XmlTextReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Fallback")
                {
                    XmlStreamHelper.ReadUntil(reader, "Fallback");
                }
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "oleItem")
                {
                    OleItems.Add(new ExcelExternalOleItem()
                    {
                        Name = reader.GetAttribute("name"),
                        Advise = XmlHelper.GetBoolFromString(reader.GetAttribute("advise")),
                        Icon = XmlHelper.GetBoolFromString(reader.GetAttribute("icon")),
                        PreferPicture = XmlHelper.GetBoolFromString(reader.GetAttribute("preferPic")),
                    });
                }
            }
        }
        internal XmlDocument ExternalOleXml;

        /// <summary>
        /// The type of external link.
        /// </summary>
        public override eExternalLinkType ExternalLinkType
        {
            get
            {
                return eExternalLinkType.OleLink;
            }
        }
        internal ZipPackageRelationship Relation
        {
            get;
            set;
        }
        /// <summary>
        /// A collection of OLE items
        /// </summary>
        public ExcelExternalOleItemsCollection OleItems
        {
            get;
        } = new ExcelExternalOleItemsCollection();
        /// <summary>
        /// The id for the connection. This is the ProgID of the OLE object
        /// </summary>
        public string ProgId { get; }

        internal override void Save(StreamWriter sw)
        {
            sw.Write($"<oleLink progId=\"{ProgId}\" r:id=\"{Relation.Id}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><oleItems>");
            foreach (ExcelExternalOleItem item in OleItems)
            {
                sw.Write(string.Format("<mc:AlternateContent><mc:Choice Requires=\"x14\"><x14:oleItem name=\"{0}\" {1}{2}{3}/></mc:Choice><mc:Fallback><oleItem name=\"{0}\" {1}{2}{3}/></mc:Fallback></mc:AlternateContent>",
                  item.Name,
                  item.Advise.GetXmlAttributeValue("advise", false),
                  item.Icon.GetXmlAttributeValue("icon", false),
                  item.PreferPicture.GetXmlAttributeValue("preferPic", false)));
            }
            sw.Write("</oleItems></oleLink>");
        }
    }
}
