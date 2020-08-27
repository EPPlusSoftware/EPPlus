/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Security.Principal;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicerCacheData : XmlHelper
    {
        const string _topPath = "x14:data/x14:tabular";
        private readonly ExcelPivotTableSlicer _slicer;
        internal ExcelPivotTableSlicerCacheData(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTableSlicer slicer) : base(nsm, topNode)
        {
            SchemaNodeOrder = new string[] { "pivotTables", "data" };
            _slicer = slicer;
        }
        const string _crossFilterPath = _topPath + "/@crossFilter";
        /// <summary>
        /// How the items that are used in slicer cross filtering are displayed
        /// </summary>
        public eCrossFilter CrossFilter
        {
            get
            {
                return GetXmlNodeString(_crossFilterPath).ToEnum(eCrossFilter.None);
            }
            set
            {
                SetXmlNodeString(_crossFilterPath, value.ToEnumString());
            }
        }

        const string _sortOrderPath = _topPath + "/@sortOrder";
        /// <summary>
        /// How the table slicer items are sorted
        /// </summary>
        public eSortOrder SortOrder
        {
            get
            {
                return GetXmlNodeString(_sortOrderPath).ToEnum(eSortOrder.None);
            }
            set
            {
                if (value == eSortOrder.None)
                {
                    DeleteNode(_sortOrderPath);
                }
                else
                {
                    SetXmlNodeString(_sortOrderPath, value.ToEnumString());
                }
            }
        }
        const string _customListSortPath = _topPath + "/@customList";

        internal void Refresh()
        {
            //Items.Clear();
            //var tbls = _field._table.CacheDefinition._cacheReference._pivotTables;
            //foreach (var item in _field.Items)
            //{
            //    var items=new pivotTableItems[]
            //    foreach (var pt in tbls)
            //    {

            //    }
            //    Items.AddInternal(new ExcelPivotTableSlicerItem())
            //}
        }

        /// <summary>
        /// If custom lists are used when sorting the items
        /// </summary>
        public bool CustomListSort
        {
            get
            {
                return GetXmlNodeBool(_customListSortPath, true);
            }
            set
            {
                SetXmlNodeBool(_customListSortPath, value, true);
            }
        }
        const string _showMissingPath = _topPath + "/@showMissing";
        /// <summary>
        /// If the source pivottable has been deleted.
        /// </summary>
        internal bool ShowMissing
        {
            get
            {
                return GetXmlNodeBool(_showMissingPath, true);
            }
            set
            {
                SetXmlNodeBool(_showMissingPath, value, true);
            }
        }
        private ExcelPivotTableSlicerItemCollection _items =null;
        public ExcelPivotTableSlicerItemCollection Items
        {
            get
            {
                if(_items==null)
                {
                    _items = new ExcelPivotTableSlicerItemCollection(_slicer);
                }
                return _items;
            }
        }

        internal void UpdateItemsXml()
        {
            var sb = new StringBuilder();
            int x = 0;
            foreach (var item in _slicer._field.Items)
            {
                if (item.Type == eItemType.Data)
                {
                    if (item.Hidden)
                        sb.Append($"<i x=\"{x++}\" />");
                    else
                        sb.Append($"<i x=\"{x++}\" s=\"1\"/>");
                }
            }

            SetXmlNodeInt(_topPath + "/@pivotCacheId", _slicer._field._table.CacheId);
            var dataNode = (XmlElement)CreateNode(_topPath+"/x14:items");
            dataNode.SetAttribute("count", x.ToString());
            dataNode.InnerXml = sb.ToString();
        }
    }
}