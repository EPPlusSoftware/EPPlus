/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer.Style
{
    /// <summary>
    /// A named table style that applies to tables only
    /// </summary>
    public class ExcelSlicerNamedStyle : XmlHelper
    {
        ExcelStyles _styles;
        internal Dictionary<eSlicerStyleElement, ExcelSlicerStyleElement> _dicSlicer = new Dictionary<eSlicerStyleElement, ExcelSlicerStyleElement>();
        internal Dictionary<eTableStyleElement, ExcelSlicerTableStyleElement> _dicTable = new Dictionary<eTableStyleElement, ExcelSlicerTableStyleElement>();
        XmlNode _tableStyleNode;
        internal ExcelSlicerNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, XmlNode tableStyleNode, ExcelStyles styles) : base(nameSpaceManager, topNode)
        {
            _styles = styles;
            if (tableStyleNode == null)
            {
                //TODO: Create table styles node with 
            }
            else
            {
                _tableStyleNode = tableStyleNode;
                foreach (XmlNode node in tableStyleNode.ChildNodes)
                {
                    if (node is XmlElement e)
                    {
                        var type = e.GetAttribute("type").ToEnum(eTableStyleElement.WholeTable);
                        _dicTable.Add(type, new ExcelSlicerTableStyleElement(nameSpaceManager, node, styles, type));
                    }
                }
            }
            if (topNode.HasChildNodes)
            {
                foreach (XmlNode node in topNode?.FirstChild?.ChildNodes)
                {
                    if (node is XmlElement e)
                    {
                        var type = e.GetAttribute("type").ToEnum(eSlicerStyleElement.SelectedItemWithData);
                        _dicSlicer.Add(type, new ExcelSlicerStyleElement(nameSpaceManager, node, styles, type));
                    }
                }
            }
        }
        private ExcelSlicerTableStyleElement GetTableStyleElement(eTableStyleElement element)
        {
            if (_dicTable.ContainsKey(element))
            {
                return _dicTable[element];
            }
            ExcelSlicerTableStyleElement item;
            item = new ExcelSlicerTableStyleElement(NameSpaceManager, _tableStyleNode, _styles, element);
            _dicTable.Add(element, item);
            return item;
        }
        private ExcelSlicerStyleElement GetSlicerStyleElement(eSlicerStyleElement element)
        {
            if (_dicSlicer.ContainsKey(element))
            {
                return _dicSlicer[element];
            }
            ExcelSlicerStyleElement item;
            item = new ExcelSlicerStyleElement(NameSpaceManager, TopNode, _styles, element);
            _dicSlicer.Add(element, item);
            return item;
        }

        /// <summary>
        /// The name of the table named style
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                if (_styles.SlicerStyles.ExistsKey(value) || _styles.TableStyles.ExistsKey(value))
                {
                    throw new InvalidOperationException("Name already exists in the collection");
                }
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Applies to the entire content of a table or pivot table
        /// </summary>
        public ExcelSlicerTableStyleElement WholeTable
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.WholeTable);
            }
        }
        /// <summary>
        /// Applies to the header row of a table or pivot table
        /// </summary>
        public ExcelSlicerTableStyleElement HeaderRow
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.HeaderRow);
            }
        }
        /// <summary>
        /// Applies to slicer item that is selected
        /// </summary>
        public ExcelSlicerStyleElement SelectedItemWithData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.SelectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a select slicer item with no data.
        /// </summary>
        public ExcelSlicerStyleElement SelectedItemWithNoData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.SelectedItemWithNoData);
            }
        }

        /// <summary>
        /// Applies to a slicer item with data that is not selected
        /// </summary>
        public ExcelSlicerStyleElement UnselectedItemWithData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.UnselectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a slicer item with no data that is not selected
        /// </summary>
        public ExcelSlicerStyleElement UnselectedItemWithNoData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.UnselectedItemWithNoData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredSelectedItemWithData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.HoveredSelectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with no data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredSelectedItemWithNoData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.HoveredSelectedItemWithNoData);
            }
        }

        /// <summary>
        /// Applies to a slicer item with data that is not selected and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredUnselectedItemWithData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.HoveredUnselectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with no data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredUnselectedItemWithNoData
        {
            get
            {
                return GetSlicerStyleElement(eSlicerStyleElement.HoveredUnselectedItemWithNoData);
            }
        }

        internal void SetFromTemplate(ExcelSlicerNamedStyle templateStyle)
        {
            foreach (var s in templateStyle._dicTable.Values)
            {
                var element = GetTableStyleElement(s.Type);
                element.Style = (ExcelDxfStyle)s.Style.Clone();
            }
            foreach (var s in templateStyle._dicSlicer.Values)
            {
                var element = GetSlicerStyleElement(s.Type);
                element.Style = (ExcelDxfStyle)s.Style.Clone();
            }
        }
        internal void SetFromTemplate(eSlicerStyle templateStyle)
        {
            LoadTableTemplate("SlicerStyles", templateStyle.ToString());
        }
        private void LoadTableTemplate(string folder, string styleName)
        {
            var zipStream = ZipHelper.OpenZipResource();
            ZipEntry entry;
            while ((entry = zipStream.GetNextEntry()) != null)
            {
                if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0 || !entry.FileName.StartsWith(folder)) continue;

                var name = new FileInfo(entry.FileName).Name;
                name = name.Substring(0, name.Length - 4);
                if (name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                {
                    var xmlContent = ZipHelper.UncompressEntry(zipStream, entry);
                    var xml = new XmlDocument();
                    xml.LoadXml(xmlContent);

                    foreach (XmlElement elem in xml.DocumentElement.ChildNodes)
                    {
                        var dxfXml = elem.InnerXml;
                        var tblType = elem.GetAttribute("name").ToEnum<eTableStyleElement>();
                        if(tblType==null)
                        {
                            var slicerType= elem.GetAttribute("name").ToEnum<eSlicerStyleElement>();
                            if(slicerType.HasValue)
                            {
                                var se = GetSlicerStyleElement(slicerType.Value);
                                var dxf = new ExcelDxfStyle(NameSpaceManager, elem.FirstChild, _styles);
                                se.Style = dxf;
                            }
                        }
                        else
                        {
                            var te = GetTableStyleElement(tblType.Value);
                            var dxf = new ExcelDxfStyle(NameSpaceManager, elem.FirstChild, _styles);
                            te.Style = dxf;
                        }
                    }
                }
            }
        }
    }
}
