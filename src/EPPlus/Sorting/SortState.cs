/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sorting
{
    /// <summary>
    /// Preserves the AutoFilter sort state.
    /// </summary>
    public class SortState : XmlHelper
    {
        internal SortState(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            _sortConditions = new SortConditionCollection(nameSpaceManager, topNode);
        }

        internal SortState(XmlNamespaceManager nameSpaceManager, ExcelWorksheet worksheet) : base(nameSpaceManager, null)
        {
            SchemaNodeOrder = worksheet.SchemaNodeOrder;
            TopNode = worksheet.WorksheetXml.SelectSingleNode(_sortStatePath, nameSpaceManager);
            if(TopNode == null)
            {
                TopNode = CreateNode(worksheet.WorksheetXml.DocumentElement, _sortStatePath);
                var attr = worksheet.WorksheetXml.CreateAttribute("xmlns:xlrd2");
                attr.Value = ExcelPackage.schemaRichData2;
                TopNode.Attributes.Append(attr);
            }
            _sortConditions = new SortConditionCollection(nameSpaceManager, TopNode);
        }

        internal SortState(XmlNamespaceManager nameSpaceManager, ExcelTable table) : base(nameSpaceManager, null)
        {
            SchemaNodeOrder = table.SchemaNodeOrder;
            TopNode = table.TableXml.SelectSingleNode(_sortStatePath, nameSpaceManager);
            if (TopNode == null)
            {
                TopNode = CreateNode(table.TableXml.DocumentElement, _sortStatePath);
                var attr = table.TableXml.CreateAttribute("xmlns:xlrd2");
                attr.Value = ExcelPackage.schemaRichData2;
                TopNode.Attributes.Append(attr);
            }
            _sortConditions = new SortConditionCollection(nameSpaceManager, TopNode);
        }

        private string _sortStatePath = "//d:sortState";
        private string _caseSensitivePath = "@caseSensitive";
        private string _columnSortPath = "@columnSort";
        private string _refPath = "@ref";

        private readonly SortConditionCollection _sortConditions;
        public SortConditionCollection SortConditions
        {
            get
            {
                return _sortConditions;
            }
        }

        /// <summary>
        /// Indicates whether or not the sort is case-sensitive
        /// </summary>
        public bool CaseSensitive
        {
            get
            {
                return GetXmlNodeBool(_caseSensitivePath);
            }
            internal set
            {
                SetXmlNodeBool(_caseSensitivePath, value, false);
            }
        }

        /// <summary>
        /// Indicates whether or not to sort by columns.
        /// </summary>
        public bool ColumnSort
        {
            get
            {
                return GetXmlNodeBool(_columnSortPath);
            }
            internal set
            {
                SetXmlNodeBool(_columnSortPath, value, false);
            }
        }

        /// <summary>
        /// The whole range of data to sort (not only the sort-by column)
        /// </summary>
        public string Ref
        {
            get
            {
                return GetXmlNodeString(_refPath);
            }
            internal set
            {
                SetXmlNodeString(_refPath, value);
            }
        }
    }
}
