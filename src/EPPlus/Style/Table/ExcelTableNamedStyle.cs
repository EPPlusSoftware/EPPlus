/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using Ionic.Zip;
using OfficeOpenXml.Core;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.IO;
using System.Reflection;
using System.Xml;

namespace OfficeOpenXml.Style.Table
{
    /// <summary>
    /// A named table style that applies to tables only
    /// </summary>
    public class ExcelTableNamedStyle : ExcelTableNamedStyleBase
    {
        internal ExcelTableNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(nameSpaceManager, topNode, styles)
        {
        }
        /// <summary>
        /// Applies to the last header cell of a table
        /// </summary>
        public ExcelTableStyleElement LastHeaderCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastHeaderCell);
            }
        }
        /// <summary>
        /// Applies to the first total cell of a table
        /// </summary>
        public ExcelTableStyleElement FirstTotalCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.FirstTotalCell);
            }
        }
        /// <summary>
        /// Applies to the last total cell of a table
        /// </summary>
        public ExcelTableStyleElement LastTotalCell
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.LastTotalCell);
            }
        }
        /// <summary>
        /// If the style applies to tables, pivot table or both
        /// </summary>
        public override eTableNamedStyleAppliesTo AppliesTo
        {
            get
            {
                return eTableNamedStyleAppliesTo.Tables;
            }
        }

        /// <summary>
        /// If a table style is applied for a table/pivot table or both
        /// </summary>
        public override eTableNamedStyleType TableNamedStyleType
        {
            get
            {
                return eTableNamedStyleType.Table;
            }
        }

        internal void SetFromTemplate(TableStyles templateStyle)
        {
            var zipStream = ZipHelper.OpenZipResource();
            ZipEntry entry;
            var ts = templateStyle.ToString();
            while ((entry = zipStream.GetNextEntry()) != null)
            {
                if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0) continue;

                var name = new FileInfo(entry.FileName).Name;
                name = name.Substring(0, name.Length - 4);
                if (name.Equals(templateStyle.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    var xmlContent = ZipHelper.UncompressEntry(zipStream, entry);
                    var xml = new XmlDocument();
                    xml.LoadXml(xmlContent);

                    foreach(XmlElement elem in xml.DocumentElement.ChildNodes)
                    {
                        var type = elem.GetAttribute("name").ToEnum(eTableStyleElement.WholeTable);
                        var dxfXml = elem.InnerXml;
                        var dxf = new ExcelDxfStyleLimitedFont(NameSpaceManager, elem.FirstChild, _styles);
                        
                        var te = GetTableStyleElement(type);
                        te.Style = dxf;
                    }
                }
            } 
        }
    }
}
