/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelSlicerPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        ExcelPivotTableSlicerCache _slicerCache;
        public ExcelSlicerPivotTableCollection(ExcelPivotTableSlicerCache slicerCache)
        {
            _slicerCache = slicerCache;
        }
        List<ExcelPivotTable> _list=new List<ExcelPivotTable>();
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public ExcelPivotTable this[int index]
        {
            get
            {
                return _list[index];
            }
        }

        internal void Add(ExcelPivotTable table)
        {
            if(_list.Count > 0 && _list[0].CacheId != table.CacheId)
            {
                throw (new InvalidOperationException("Multiple Pivot tables added to a slicer must refer to the same cache."));
            }
            _list.Add(table);
            _slicerCache.UpdateItemsXml();
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
    }
    public class ExcelPivotTableSlicerCache : ExcelSlicerCache
    {
        internal ExcelPivotTableSlicerCache(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            PivotTables = new ExcelSlicerPivotTableCollection(this);
        }

        internal ExcelPivotTableSlicer _slicer;
        internal void Init(ExcelWorkbook wb, string name, ExcelPivotTableSlicer slicer)
        {
            CreatePart(wb);
            TopNode = SlicerCacheXml.DocumentElement;
            Name = "Slicer_" + name;
            SourceName = slicer._field.Cache.Name;
            _slicer = slicer;
            wb.Names.AddFormula(Name, "#N/A");
            PivotTables.Add(slicer._field._table);
            CreateWorkbookReference(wb, ExtLstUris.WorkbookSlicerPivotTableUri);

            Data.Items.Refresh();
        }
        /// <summary>
        /// Init must be called before accessing any properties as it sets several properties.
        /// </summary>
        /// <param name="wb"></param>
        internal override void Init(ExcelWorkbook wb)
        {
            foreach (XmlElement ptElement in GetNodes("x14:pivotTables/x14:pivotTable"))
            {
                var name = ptElement.GetAttribute("name");
                var tabId = ptElement.GetAttribute("tabId");

                if(int.TryParse(tabId, out int sheetId))
                {
                    var ws = wb.Worksheets.GetBySheetID(sheetId);
                    var pt = ws?.PivotTables[name];
                    if(pt!=null)
                    {
                        PivotTables.Add(pt);
                    }
                }
            }
        }
        internal void Init(ExcelWorkbook wb, ExcelPivotTableSlicer slicer)
        {
            _slicer = slicer;
            Init(wb);
            _slicer._field = PivotTables[0].Fields[SourceName];
        }
        public override eSlicerSourceType SourceType
        {
            get
            {
                return eSlicerSourceType.PivotTable;
            }   
        }
        public ExcelSlicerPivotTableCollection PivotTables { get; }
        ExcelPivotTableSlicerCacheTabularData _data=null;
        public ExcelPivotTableSlicerCacheTabularData Data 
        { 
            get
            {
                if(_data==null)
                {
                    _data = new ExcelPivotTableSlicerCacheTabularData(NameSpaceManager, TopNode, _slicer);
                }
                return _data;
            }
        }

        protected internal void UpdateItemsXml()
        {
           var sb = new StringBuilder();
            foreach(var pt in PivotTables)
            {
                sb.Append($"<pivotTable name=\"{pt.Name}\" tabId=\"{pt.WorkSheet.SheetId}\"/>");
            }
            var ptNode = CreateNode("x14:pivotTables");
            ptNode.InnerXml = sb.ToString();
            Data.UpdateItemsXml();
        }
    }
}
