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
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelSlicerPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        List<ExcelPivotTable> _list=new List<ExcelPivotTable>();
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        internal void Add(ExcelPivotTable table)
        {
            _list.Add(table);
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

        }
        /// <summary>
        /// Init must be called before accessing any properties as it sets several properties.
        /// </summary>
        /// <param name="wb"></param>
        internal override void Init(ExcelWorkbook wb)
        {            
            foreach(XmlElement ptElement in GetNodes("x14:pivotTables/x14:pivotTable"))
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
        public override eSlicerSourceType SourceType
        {
            get
            {
                return eSlicerSourceType.PivotTable;
            }   
        }
        public ExcelSlicerPivotTableCollection PivotTables { get; } = new ExcelSlicerPivotTableCollection();
        ExcelPivotTableSlicerCacheData _data=null;
        public ExcelPivotTableSlicerCacheData Data 
        { 
            get
            {
                if(_data==null)
                {
                    _data = new ExcelPivotTableSlicerCacheData(NameSpaceManager, TopNode);
                }
                return _data;
            }
        }
    }
}
