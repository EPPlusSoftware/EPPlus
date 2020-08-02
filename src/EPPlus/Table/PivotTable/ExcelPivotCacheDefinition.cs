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
using System.Linq;
using OfficeOpenXml.Utils;
using System.Security;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Cache definition. This class defines the source data. Note that one cache definition can be shared between many pivot tables.
    /// </summary>
    public class ExcelPivotCacheDefinition : XmlHelper
    {
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable) :
            base(ns, null)
        {
            foreach (var r in pivotTable.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition"))
            {
                Relationship = r;
            }
            CacheDefinitionUri = UriHelper.ResolvePartUri(Relationship.SourceUri, Relationship.TargetUri);

            var pck = pivotTable.WorkSheet._package.ZipPackage;
            Part = pck.GetPart(CacheDefinitionUri);
            CacheDefinitionXml = new XmlDocument();
            LoadXmlSafe(CacheDefinitionXml, Part.GetStream());

            TopNode = CacheDefinitionXml.DocumentElement;
            PivotTable = pivotTable;
        }
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tblId) :
            base(ns, null)
        {
            PivotTable = pivotTable;

            var pck = pivotTable.WorkSheet._package.ZipPackage;
            
            //CacheDefinition
            CacheDefinitionXml = new XmlDocument();
            LoadXmlSafe(CacheDefinitionXml, GetStartXml(sourceAddress), Encoding.UTF8);
            CacheDefinitionUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref tblId); 
            Part = pck.CreatePart(CacheDefinitionUri, ExcelPackage.schemaPivotCacheDefinition);
            TopNode = CacheDefinitionXml.DocumentElement;

            //CacheRecord. Create an empty one.
            CacheRecordUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheRecords{0}.xml", ref tblId); 
            var cacheRecord = new XmlDocument();
            cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
            var recPart = pck.CreatePart(CacheRecordUri, ExcelPackage.schemaPivotCacheRecords);
            cacheRecord.Save(recPart.GetStream());

            RecordRelationship = Part.CreateRelationship(UriHelper.ResolvePartUri(CacheDefinitionUri, CacheRecordUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
            RecordRelationshipID = RecordRelationship.Id;

            CacheDefinitionXml.Save(Part.GetStream());
        }        
        /// <summary>
        /// Reference to the internal package part
        /// </summary>
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the cache definition in the package.
        /// </summary>
        public XmlDocument CacheDefinitionXml { get; private set; }
        /// <summary>
        /// The package internal URI to the pivottable cache definition Xml Document.
        /// </summary>
        public Uri CacheDefinitionUri
        {
            get;
            internal set;
        }
        internal Uri CacheRecordUri
        {
            get;
            set;
        }
        internal Packaging.ZipPackageRelationship Relationship
        {
            get;
            set;
        }
        internal Packaging.ZipPackageRelationship RecordRelationship
        {
            get;
            set;
        }
        internal string RecordRelationshipID 
        {
            get
            {
                return GetXmlNodeString("@r:id");
            }
            set
            {
                SetXmlNodeString("@r:id", value);
            }
        }
        /// <summary>
        /// Referece to the PivoTable object
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
                    if (CacheSource == eSourceType.Worksheet)
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

                SetXmlNodeString(_sourceWorksheetPath, value.Worksheet.Name);
                SetXmlNodeString(_sourceAddressPath, value.FirstAddress);
                _sourceRange = value;
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
                var s=GetXmlNodeString("d:cacheSource/@type");
                if (s == "")
                {
                    return eSourceType.Worksheet;
                }
                else
                {
                    return (eSourceType)Enum.Parse(typeof(eSourceType), s, true);
                }
            }
        }
        private string GetStartXml(ExcelRangeBase sourceAddress)
        {
            string xml="<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

            xml += "<cacheSource type=\"worksheet\">";
            xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceAddress.Address, sourceAddress.WorkSheetName);
            xml += "</cacheSource>";
            xml += string.Format("<cacheFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
            var sourceWorksheet = PivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheetName];
            for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
            {
                var name = sourceWorksheet?.GetValueInner(sourceAddress._fromRow, col);
                if (name==null || name.ToString()=="")
                {
                    xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceAddress._fromCol + 1);
                }
                else
                {
                    xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", SecurityElement.Escape(name.ToString()));
                }
                xml += "<sharedItems containsBlank=\"1\" /> ";
                xml += "</cacheField>";
            }
            xml += "</cacheFields>";
            xml += "</pivotCacheDefinition>";

            return xml;
        }
    }
}
