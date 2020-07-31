/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using System.Security;
using System.Linq;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Cache definition. This class defines the source data. Note that one cache definition can be shared between many pivot tables.
    /// </summary>
    public class ExcelPivotCacheDefinition : XmlHelper
    {
        ExcelWorkbook _wb;
        private PivotTableCacheInternal _cacheReference;
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable) :
            base(ns, null)
        {
            var cacheDefinitionUri=pivotTable.GetCacheUriFromRel();
            _wb=PivotTable.WorkSheet.Workbook;
            var c = _wb._pivotTableCaches.Values.FirstOrDefault(x => x.CacheDefinitionUri.OriginalString == cacheDefinitionUri.OriginalString);
            if (c == null)
            {
                var pck = pivotTable.WorkSheet._package.Package;
                _cacheReference = new PivotTableCacheInternal(ns)
                {
                    Part = pck.GetPart(cacheDefinitionUri),
                    CacheDefinitionUri = CacheDefinitionUri,
                    CacheDefinitionXml = new XmlDocument(),
                };
                _cacheReference.Init();
            }
            else
            {
                _cacheReference = c;
            }
            PivotTable = pivotTable;
        }
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tblId) :
            base(ns, null)
        {
            PivotTable = pivotTable;
            _wb = PivotTable.WorkSheet.Workbook;
            if (_wb._pivotTableCaches.TryGetValue(sourceAddress.FullAddress, out _cacheReference))
            {
                _cacheReference._pivotTables.Add(pivotTable);
            }
            else
            {
                _cacheReference = new PivotTableCacheInternal(ns);
                _cacheReference.InitNew(pivotTable, sourceAddress, tblId, null);
                _wb._pivotTableCaches.Add(sourceAddress.FullAddress, _cacheReference);
            }
        }        
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the cache definition in the package.
        /// </summary>
        public XmlDocument CacheDefinitionXml
        { 
            get
            {
                return _cacheReference.CacheDefinitionXml;
            }
        }
        /// <summary>
        /// The package internal URI to the pivottable cache definition Xml Document.
        /// </summary>
        public Uri CacheDefinitionUri
        {
            get
            {
                return _cacheReference.CacheDefinitionUri;
            }
        }
        internal ZipPackageRelationship Relationship 
        { 
            get
            {
                return _cacheReference.Relationship;
            }
        }
        /// <summary>
        /// Referece to the PivotTable object
        /// </summary>
        public ExcelPivotTable PivotTable
        {
            get;
            private set;
        }
        
        const string _sourceWorksheetPath="d:cacheSource/d:worksheetSource/@sheet";
        internal const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";
        internal const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";
        internal ExcelRangeBase _sourceRange = null;
        /// <summary>
        /// The source data range when the pivottable has a worksheet datasource. 
        /// The number of columns in the range must be intact if this property is changed.
        /// The range must be in the same workbook as the pivottable.
        /// </summary>
        public ExcelRangeBase SourceRange
        {
            get
            {
                if (_sourceRange == null)
                {
                    if (_cacheReference.CacheSource == eSourceType.Worksheet)
                    {
                        var ws = PivotTable.WorkSheet.Workbook.Worksheets[GetXmlNodeString(_sourceWorksheetPath)];
                        if (ws == null) //Not worksheet, check name or table name
                        {
                            var name = GetXmlNodeString(_sourceNamePath);
                            foreach (var n in PivotTable.WorkSheet.Workbook.Names)
                            {
                                if(name.Equals(n.Name,StringComparison.OrdinalIgnoreCase))
                                {
                                    _sourceRange = n;
                                    return _sourceRange;
                                }
                            }
                            foreach (var w in PivotTable.WorkSheet.Workbook.Worksheets)
                            {
                                _sourceRange = GetRangeByName(w, name);
                                if (_sourceRange != null) break;
                            }
                        }
                        else
                        {
                            var address = GetXmlNodeString(_sourceAddressPath);
                            if(string.IsNullOrEmpty(address))
                            {
                                var name = GetXmlNodeString(_sourceNamePath);
                                _sourceRange = GetRangeByName(ws, name);
                            }
                            else
                            {
                                _sourceRange = ws.Cells[address];

                            }
                        }
                    }
                    else
                    {
                        throw (new ArgumentException("The cachesource is not a worksheet"));
                    }
                }
                return _sourceRange;
            }
            set
            {
                if (PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
                {
                    throw (new ArgumentException("Range must be in the same package as the pivottable"));
                }

                var sr=SourceRange;
                if (value.End.Column - value.Start.Column != sr.End.Column - sr.Start.Column)
                {
                    throw (new ArgumentException("Can not change the number of columns(fields) in the SourceRange"));
                }
                if (value.FullAddress == SourceRange.FullAddress) return; //Same
                if(_wb._pivotTableCaches.TryGetValue(value.FullAddress, out PivotTableCacheInternal cache))
                {
                    _cacheReference._pivotTables.Remove(PivotTable);
                    cache._pivotTables.Add(PivotTable);
                    _cacheReference = cache;
                }
                else if(_cacheReference._pivotTables.Count==1)
                {
                    SetXmlNodeString(_sourceWorksheetPath, value.Worksheet.Name);
                    SetXmlNodeString(_sourceAddressPath, value.FirstAddress);
                    _sourceRange = value;
                }
                else
                {
                    _cacheReference._pivotTables.Remove(PivotTable);
                    var xml = _cacheReference.CacheDefinitionXml;
                    _cacheReference = new PivotTableCacheInternal(NameSpaceManager);
                    _cacheReference.InitNew(PivotTable, value, _wb._nextPivotTableID++, xml);                    
                }
            }
        }

        private ExcelRangeBase GetRangeByName(ExcelWorksheet w, string name)
        {
            if (w is ExcelChartsheet) return null;
            if (w.Tables._tableNames.ContainsKey(name))
            {
                return w.Cells[w.Tables[name].Address.Address];
            }
            foreach (var n in w.Names)
            {
                if (name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return n;
                }
            }
            return null;
        }

        /// <summary>
        /// Type of source data
        /// </summary>
        public eSourceType CacheSource
        {
            get
            {
                return _cacheReference.CacheSource;
            }
        }
    }
}
