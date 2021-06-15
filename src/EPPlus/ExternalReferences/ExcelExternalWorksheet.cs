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
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using System;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// A representation of an external cached worksheet.
    /// </summary>
    public class ExcelExternalWorksheet : IExcelExternalNamedItem
    {
        internal ExcelExternalWorksheet()
        {
            CachedNames = new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>();
            CellValues = new ExcelExternalCellCollection(new CellStore<object>(), new CellStore<int>());
        }

        internal ExcelExternalWorksheet(
            CellStore<object> values,
            CellStore<int> metaData,
            ExcelExternalNamedItemCollection<ExcelExternalDefinedName> definedNames)
        {
            CachedNames = definedNames;
            CellValues = new ExcelExternalCellCollection(values, metaData);
        }
        /// <summary>
        /// The sheet id
        /// </summary>
        public int SheetId { get; internal set; }
        /// <summary>
        /// The name of the worksheet.
        /// </summary>
        public string Name { get; internal set; }
        /// <summary>
        /// If errors have occured on the last update of the cached values.
        /// </summary>
        public bool RefreshError { get; internal set; }
        /// <summary>
        /// A collection of cached names for an external worksheet
        /// </summary>
        public ExcelExternalNamedItemCollection<ExcelExternalDefinedName> CachedNames { get; }
        /// <summary>
        /// Cached cell values for the worksheet. Only cells referenced in the workbook are stored in the cache.
        /// </summary>
        public ExcelExternalCellCollection CellValues 
        { 
            get; 
        }
        public override string ToString()
        {
            return Name;
        }
    }
}
