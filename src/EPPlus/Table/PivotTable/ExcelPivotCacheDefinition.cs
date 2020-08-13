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
    public class ExcelPivotCacheDefinition
    {
        ExcelWorkbook _wb;
        internal PivotTableCacheInternal _cacheReference;
        XmlNamespaceManager _nsm;
        internal ExcelPivotCacheDefinition(XmlNamespaceManager nsm, ExcelPivotTable pivotTable)
        {
            var cacheDefinitionUri = pivotTable.GetCacheUriFromRel();
            _wb = PivotTable.WorkSheet.Workbook;
            _nsm = nsm;
            var c = _wb._pivotTableCaches.Values.FirstOrDefault(x => x.CacheDefinitionUri.OriginalString == cacheDefinitionUri.OriginalString);
            if (c == null)
            {
                var pck = pivotTable.WorkSheet._package.Package;
                _cacheReference = new PivotTableCacheInternal(nsm, _wb)
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
        internal ExcelPivotCacheDefinition(XmlNamespaceManager nsm, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress)
        {
            PivotTable = pivotTable;
            _wb = PivotTable.WorkSheet.Workbook;
            _nsm = nsm;
            if (_wb._pivotTableCaches.TryGetValue(sourceAddress.FullAddress, out _cacheReference))
            {
                _cacheReference._pivotTables.Add(pivotTable);
            }
            else
            {
                _cacheReference = new PivotTableCacheInternal(nsm, _wb);
                _cacheReference.InitNew(pivotTable, sourceAddress, null);
                _wb._pivotTableCaches.Add(sourceAddress.FullAddress, _cacheReference);
                _wb.AddPivotTableCache(_cacheReference.CacheId.ToString(), _cacheReference.CacheDefinitionUri);
            }
            var rel = pivotTable.Part.CreateRelationship(UriHelper.ResolvePartUri(pivotTable.PivotTableUri, _cacheReference.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
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

        const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
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
                return _cacheReference.SourceRange;
            }
            set
            {
                if (PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
                {
                    throw (new ArgumentException("Range must be in the same package as the pivottable"));
                }

                var sr = SourceRange;
                if (value.End.Column - value.Start.Column != sr.End.Column - sr.Start.Column)
                {
                    throw (new ArgumentException("Can not change the number of columns(fields) in the SourceRange"));
                }
                if (value.FullAddress == SourceRange.FullAddress) return; //Same
                if (_wb._pivotTableCaches.TryGetValue(value.FullAddress, out PivotTableCacheInternal cache))
                {
                    _cacheReference._pivotTables.Remove(PivotTable);
                    cache._pivotTables.Add(PivotTable);
                    _cacheReference = cache;
                    PivotTable.CacheId = _cacheReference.CacheId;
                }
                else if (_cacheReference._pivotTables.Count == 1)
                {
                    _cacheReference.SetXmlNodeString(_sourceWorksheetPath, value.Worksheet.Name);
                    _cacheReference.SetXmlNodeString(_sourceAddressPath, value.FirstAddress);
                    _sourceRange = value;
                }
                else
                {
                    _cacheReference._pivotTables.Remove(PivotTable);
                    var xml = _cacheReference.CacheDefinitionXml;
                    _cacheReference = new PivotTableCacheInternal(_nsm, _wb);
                    _cacheReference.InitNew(PivotTable, value, xml.InnerXml);
                    PivotTable.CacheId = _cacheReference.CacheId;
                    _wb.AddPivotTableCache(_cacheReference.CacheId.ToString(), _cacheReference.CacheDefinitionUri);
                    _wb._pivotTableCaches.Add(value.FullAddress, _cacheReference);
                }
            }
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
