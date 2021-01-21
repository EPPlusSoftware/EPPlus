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
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
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
        internal Dictionary<eTableStyleElement, ExcelTableStyleElement> _dicTable = new Dictionary<eTableStyleElement, ExcelTableStyleElement>();
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
                        _dicTable.Add(type, new ExcelTableStyleElement(nameSpaceManager, node, styles, type));
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
        private ExcelTableStyleElement GetTableStyleElement(eTableStyleElement element)
        {
            if (_dicTable.ContainsKey(element))
            {
                return _dicTable[element];
            }
            ExcelTableStyleElement item;
            item = new ExcelTableStyleElement(NameSpaceManager, _tableStyleNode, _styles, element);
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
        public ExcelTableStyleElement WholeTable
        {
            get
            {
                return GetTableStyleElement(eTableStyleElement.WholeTable);
            }
        }
        /// <summary>
        /// Applies to the header row of a table or pivot table
        /// </summary>
        public ExcelTableStyleElement HeaderRow
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

    }
}
