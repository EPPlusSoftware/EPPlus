/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/29/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// A slicer cache with a table as source
    /// </summary>
    public class ExcelTableSlicerCache : ExcelSlicerCache
    {
        const string _extPath = "x14:extLst/d:ext";
        const string _topPath = _extPath+"/x15:tableSlicerCache";
        internal ExcelTableSlicerCache(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
        }

        internal override void Init(ExcelWorkbook wb)
        {
            var tbl = wb.GetTable(TableId);                
            TableColumn = tbl?.Columns.FirstOrDefault(x => x.Id == ColumnId);
        }
        internal void Init(ExcelTableColumn column)
        {
            var wb = column.Table.WorkSheet.Workbook;
            var p = wb._package.Package;
            var uri = GetNewUri(p, "/xl/slicerCaches/slicerCache{0}.xml");
            Part = p.CreatePart(uri, "application/vnd.ms-excel.slicerCache+xml");
            CacheRel = wb.Part.CreateRelationship(UriHelper.GetRelativeUri(wb.WorkbookUri, uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicerCache);
            SlicerCacheXml = new XmlDocument();
            SlicerCacheXml.LoadXml(GetStartXml());
            SlicerCacheXml.DocumentElement.InnerXml = $"<extLst><x:ext uri=\"{{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:tableSlicerCache tableId=\"{column.Table.Id}\" column=\"{column.Id}\"/></x:ext></extLst>";
            TopNode = SlicerCacheXml.DocumentElement;
            Name = "Slicer_" + column.Name;
            SourceName = column.Name;
            wb.Names.AddFormula(Name, "#N/A");

            var slNode = wb.GetExtLstSubNode("{46BE6895-7355-4a93-B00E-2C351335B9C9}", "x15:slicerCaches");
            if (slNode == null)
            {
                wb.CreateNode("d:extLst/d:ext", false, true);
                slNode = wb.CreateNode("d:extLst/d:ext/x15:slicerCaches", false, true);
                ((XmlElement)slNode.ParentNode).SetAttribute("uri", "{46BE6895-7355-4a93-B00E-2C351335B9C9}");
            }
            var xh = XmlHelperFactory.Create(NameSpaceManager, slNode);
            var element = (XmlElement)xh.CreateNode("x14:slicerCache", false, true);
            element.SetAttribute("id", ExcelPackage.schemaRelationships, CacheRel.Id);
        }
        public override eSlicerSourceType SourceType 
        {
            get
            {
                return eSlicerSourceType.Table;
            }
        }
        /// <summary>
        /// The table column that is the source for the slicer
        /// </summary>
        public ExcelTableColumn TableColumn
        {
            get;
            private set;
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
                if(value==eSortOrder.None)
                {
                    DeleteNode(_sortOrderPath);
                }
                else
                {
                    SetXmlNodeString(_sortOrderPath, value.ToEnumString());
                }
            }
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
        const string _customListSortPath = _topPath + "/@customList";
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
        const string _hideItemsWithNoDataPath = "x15:slicerCacheHideItemsWithNoData";
        /// <summary>
        /// If true, items that have no data are not displayed
        /// </summary>
        public bool HideItemsWithNoData 
        { 
            get
            {
                return ExistNode(_extPath +"/" + _hideItemsWithNoDataPath);
            }
            set
            {
                if(value)
                {
                    var node = CreateNode("x14:extLst", false, true);
                    var helper = XmlHelperFactory.Create(NameSpaceManager, node);
                    helper.CreateNode("d:ext/"+_hideItemsWithNoDataPath);
                }
                else
                {
                    DeleteAllNode(_extPath + "/" + _hideItemsWithNoDataPath);
                }
            }
        }
        const string _columnIndexPath = _topPath + "/@column";
        internal int ColumnId
        {
            get
            {
                return GetXmlNodeInt(_columnIndexPath);
            }
            set
            {
                SetXmlNodeInt(_columnIndexPath, value);
            }
        }
        const string _tableIdPath = _topPath + "/@tableId";
        private ExcelTableColumn column;

        internal int TableId
        {
            get
            {
                return GetXmlNodeInt(_tableIdPath);
            }
            set
            {
                SetXmlNodeInt(_tableIdPath, value);
            }
        }
    }
}
