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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// Represents a pivot table slicer cache.
    /// </summary>
    public class ExcelPivotTableSlicerCache : ExcelSlicerCache
    {
        internal ExcelPivotTableField _field=null;
        internal ExcelPivotTableSlicerCache(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            PivotTables = new ExcelSlicerPivotTableCollection(this);
        }

        internal void Init(ExcelWorkbook wb, string name, ExcelPivotTableField field)
        {
            if(wb._slicerCaches==null) 
                wb.LoadSlicerCaches();

            CreatePart(wb);
            TopNode = SlicerCacheXml.DocumentElement;
            Name = "Slicer_" + ExcelAddressUtil.GetValidName(name);
            _field = field;
            SourceName = _field.Cache.Name;
            wb.Names.AddFormula(Name, "#N/A");
            PivotTables.Add(_field.PivotTable);           
            CreateWorkbookReference(wb, ExtLstUris.WorkbookSlicerPivotTableUri);
            SlicerCacheXml.Save(Part.GetStream());
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
                        if (_field == null)
                        {
                            _field = pt.Fields.Where(x => x.Cache.Name.Equals(SourceName, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                        }

                        PivotTables._list.Add(pt);
                    }
                }
            }
        }
        internal void Init(ExcelWorkbook wb, ExcelPivotTableSlicer slicer)
        {
            _field = PivotTables[0].Fields.Where(x=>x.Cache.Name==SourceName).FirstOrDefault();
            Init(wb);
        }
        /// <summary>
        /// The source type of the slicer
        /// </summary>
        public override eSlicerSourceType SourceType
        {
            get
            {
                return eSlicerSourceType.PivotTable;
            }   
        }
        /// <summary>
        /// A collection of pivot tables attached to the slicer cache.
        /// </summary>
        public ExcelSlicerPivotTableCollection PivotTables { get; }
        ExcelPivotTableSlicerCacheTabularData _data=null;
        /// <summary>
        /// Tabular data for a pivot table slicer cache.
        /// </summary>
        public ExcelPivotTableSlicerCacheTabularData Data 
        { 
            get
            {
                if(_data==null)
                {
                    _data = new ExcelPivotTableSlicerCacheTabularData(NameSpaceManager, TopNode, this);
                }
                return _data;
            }
        }

        internal void UpdateItemsXml()
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
