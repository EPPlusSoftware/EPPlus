using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Security;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    internal class PivotTableCacheInternal : XmlHelper
    {
        internal List<ExcelPivotTable> _pivotTables=new List<ExcelPivotTable>();

        public PivotTableCacheInternal(XmlNamespaceManager nsm) : base(nsm)
        {

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
        internal XmlDocument CacheDefinitionXml { get; set; }
        /// <summary>
        /// The package internal URI to the pivottable cache definition Xml Document.
        /// </summary>
        internal Uri CacheDefinitionUri
        {
            get;
            set;
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
        internal eSourceType CacheSource
        {
            get
            {
                var s = GetXmlNodeString("d:cacheSource/@type");
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
        internal void Init()
        {
            LoadXmlSafe(CacheDefinitionXml, Part.GetStream());
            TopNode = CacheDefinitionXml.DocumentElement;
        }
        internal void InitNew(ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tblId, XmlDocument xml)
        {
            var pck = pivotTable.WorkSheet._package.Package;

            CacheDefinitionXml = new XmlDocument();
            var sourceWorksheet = pivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheetName];
            if(xml==null)
            {
                LoadXmlSafe(CacheDefinitionXml, GetStartXml(sourceWorksheet, sourceAddress), Encoding.UTF8);
            }
            else
            {
                CacheDefinitionXml = xml;

            }
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
            _pivotTables.Add(pivotTable);
        }
        private string GetStartXml(ExcelWorksheet sourceWorksheet, ExcelRangeBase sourceAddress)
        {
            string xml = "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

            xml += "<cacheSource type=\"worksheet\">";
            xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceAddress.Address, sourceAddress.WorkSheetName);
            xml += "</cacheSource>";
            xml += string.Format("<cacheFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
            for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
            {
                var name = sourceWorksheet?.GetValueInner(sourceAddress._fromRow, col);
                if (name == null || name.ToString() == "")
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
