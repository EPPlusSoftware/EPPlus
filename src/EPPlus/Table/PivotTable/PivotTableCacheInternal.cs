using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Handles the pivot table cache.
    /// </summary>
    internal class PivotTableCacheInternal : XmlHelper
    {
        internal List<ExcelPivotTable> _pivotTables=new List<ExcelPivotTable>();
        internal readonly ExcelWorkbook _wb;
        public PivotTableCacheInternal(XmlNamespaceManager nsm, ExcelWorkbook wb) : base(nsm)
        {
            _wb = wb;
        }
        public PivotTableCacheInternal(ExcelWorkbook wb, Uri uri, int cacheId) : base (wb.NameSpaceManager)
        {
            _wb = wb;
            CacheDefinitionUri = uri;
            Part = wb._package.ZipPackage.GetPart(uri);

            CacheDefinitionXml = new XmlDocument();
            LoadXmlSafe(CacheDefinitionXml, Part.GetStream());
            TopNode = CacheDefinitionXml.DocumentElement;
            CacheId = cacheId;

            ZipPackageRelationship rel = Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheRecords").FirstOrDefault();
            if (rel != null)
            {
                CacheRecordUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            }

            _wb.SetNewPivotCacheId(cacheId);
        }
        const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
        internal const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";
        internal const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";
        internal string Ref
        {
            get
            {
                return GetXmlNodeString(_sourceAddressPath);
            }
        }
        internal string SourceName
        {
            get
            {
                return GetXmlNodeString(_sourceNamePath);
            }
        }        
        internal ExcelRangeBase SourceRange 
        { 
            get
            {
                ExcelRangeBase sourceRange=null;
                if (CacheSource == eSourceType.Worksheet)
                {
                    var ws = _wb.Worksheets[GetXmlNodeString(_sourceWorksheetPath)];
                    if (ws == null) //Not worksheet, check name or table name
                    {
                        var name = GetXmlNodeString(_sourceNamePath);
                        foreach (var n in _wb.Names)
                        {
                            if (name.Equals(n.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                sourceRange = n;
                                return sourceRange;
                            }
                        }
                        foreach (var w in _wb.Worksheets)
                        {
                            sourceRange = GetRangeByName(w, name);
                            if (sourceRange != null) break;
                        }
                    }
                    else
                    {
                        var address = Ref;
                        if (string.IsNullOrEmpty(address))
                        {
                            var name = SourceName;
                            sourceRange = GetRangeByName(ws, name);
                        }
                        else
                        {
                            sourceRange = ws.Cells[address];
                        }
                    }
                }
                else
                {
                    throw (new ArgumentException("The cache source is not a worksheet"));
                }
                return sourceRange;
            }

        }
        private ExcelRangeBase GetRangeByName(ExcelWorksheet w, string name)
        {
            if (w is ExcelChartsheet) return null;
            if (w.Tables._tableNames.ContainsKey(name))
            {
                var t = w.Tables[name];
                var toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                return w.Cells[t.Address._fromRow, t.Address._fromCol, toRow, t.Address._toCol];
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
        internal Packaging.ZipPackageRelationship RecordRelationship
        {
            get;
            set;
        }
        internal string RecordRelationshipId
        {
            get
            {
                return GetXmlNodeString("@r:id");
            }
            set
            {
                SetXmlNodeString("@r:id", value, true);
            }
        }
        List<ExcelPivotTableCacheField> _fields=null;
        internal List<ExcelPivotTableCacheField> Fields
        {
            get
            {
                if(_fields == null)
                {
                    LoadFields();
                    //RefreshFields();
                }
                return _fields;
            }
        }

        private void LoadFields()
        {
            //Add fields.
            var index = 0;
            _fields = new List<ExcelPivotTableCacheField>();
            foreach (XmlNode node in CacheDefinitionXml.DocumentElement.SelectNodes("d:cacheFields/d:cacheField", NameSpaceManager))
            {
                _fields.Add(new ExcelPivotTableCacheField(NameSpaceManager, node, this, index++));
            }
        }

        internal void RefreshFields()
        {
            var tableFields = GetTableFields();
            var fields = new List<ExcelPivotTableCacheField>();
            var r = SourceRange;
            bool cacheUpdated=false;
            
            for (int col = r._fromCol; col <= r._toCol; col++)
            {
                var ix = col - r._fromCol;
                if (_fields!=null && ix < _fields.Count && _fields[ix].Grouping != null)
                {
                    fields.Add(_fields[ix]);
                }
                else
                {
                    var ws = r.Worksheet;
                    var name = ws.GetValue(r._fromRow, col)?.ToString().Trim();
                    ExcelPivotTableCacheField field;
                    if (_fields==null || ix>=_fields?.Count)
                    {
                        if (string.IsNullOrEmpty(name))
                        {
                            throw new InvalidOperationException($"Pivot Cache with id {CacheId} is invalid . Contains reference to a column with an empty header");
                        }
                        field = CreateField(name, ix);
                        field.TopNode.InnerXml = "<sharedItems/>";
                        foreach(var pt in _pivotTables)
                        {
                            pt.Fields.AddField(ix);
                        }
                        cacheUpdated = true;
                    }
                    else
                    {
                        field=_fields[ix];

                        field.SharedItems.Clear();

                        if (cacheUpdated == false && string.IsNullOrEmpty(name)==false && !field.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase)) cacheUpdated=true;
                    }

                    if (!string.IsNullOrEmpty(name) && !field.Name.StartsWith(name)) field.Name = name;
                    var hs = new HashSet<object>();
                    var dimensionToRow = ws.Dimension?._toRow ?? r._fromRow + 1;
                    var toRow = r._toRow < dimensionToRow ? r._toRow : dimensionToRow;
                    for (int row = r._fromRow + 1; row <= toRow; row++)
                    {
                        ExcelPivotTableCacheField.AddSharedItemToHashSet(hs, ws.GetValue(row, col));
                    }
                    field.SharedItems._list = hs.ToList();
                    fields.Add(field);
                }
            }
            if (_fields != null)
            {
                for (int i = fields.Count; i < _fields.Count; i++)
                {
                    fields.Add(_fields[i]);
                }
            }
            _fields = fields;

            if(cacheUpdated) UpdateRowColumnPageFields(tableFields);

             RefreshPivotTableItems();
        }

        private void UpdateRowColumnPageFields(List<List<string>> tableFields)
        {
            for(int tblIx=0;tblIx<_pivotTables.Count;tblIx++)
            {

                var l = tableFields[tblIx];
                var tbl = _pivotTables[tblIx];
                tbl.PageFields._list.ForEach(x => { x.IsPageField = false; x.Axis = ePivotFieldAxis.None; });
                tbl.ColumnFields._list.ForEach(x => { x.IsColumnField = false; x.Axis = ePivotFieldAxis.None; });
                tbl.RowFields._list.ForEach(x => { x.IsRowField = false; x.Axis = ePivotFieldAxis.None; });
                tbl.DataFields._list.ForEach(x => { x.Field.IsDataField = false; x.Field.Axis = ePivotFieldAxis.None; });

                ChangeIndex(tbl.PageFields, l);
                ChangeIndex(tbl.ColumnFields, l);
                ChangeIndex(tbl.RowFields, l);
                for (int i = 0; i < tbl.DataFields.Count; i++)
                {
                    var df = tbl.DataFields[i];
                    var prevName = l[df.Index];
                    var newIx = _fields.FindIndex(x => x.Name.Equals(prevName, StringComparison.CurrentCultureIgnoreCase));
                    if (newIx >= 0)
                    {
                        df.Index = newIx;
                        df.Field = tbl.Fields[newIx];
                        df.Field.IsDataField = true;
                    }
                    else
                    {
                        tbl.DataFields._list.RemoveAt(i--);
                    }

                    foreach (ExcelPivotTableAreaStyle s in tbl.Styles)
                    {
                        if (s.FieldIndex == df.Index)
                        {
                            s.FieldIndex = newIx;
                        }
                        foreach (ExcelPivotAreaReference c in s.Conditions.Fields)
                        {
                            if (c.FieldIndex == df.Index)
                            {
                                c.FieldIndex = newIx;
                            }
                        }
                        
                        if (s.Conditions.DataFields.FieldIndex == df.Index)
                        {
                            s.Conditions.DataFields.FieldIndex = newIx;
                        }
                    }
                }
            }
        }

        private void ChangeIndex(ExcelPivotTableRowColumnFieldCollection fields, List<string> prevFields)
        {
            var newFields = new List<ExcelPivotTableField>();
            for (int i = 0; i < fields.Count; i++)
            {
                var f = fields[i];
                var prevName = prevFields[f.Index];
                var ix = _fields.FindIndex(x => x.Name.Equals(prevName, StringComparison.CurrentCultureIgnoreCase));
                if (ix>=0)
                {
                    var fld = fields._table.Fields[ix];

                    newFields.Add(fld);
                    if(fld.PageFieldSettings!=null)
                    {
                        fld.PageFieldSettings.Index = ix;
                        fld.PageFieldSettings._field = fld;
                    }
                    foreach(ExcelPivotTableAreaStyle s in f._pivotTable.Styles)
                    {
                        if(s.FieldIndex==f.Index)
                        {
                            s.FieldIndex = ix;
                        }
                        foreach(ExcelPivotAreaReference c in s.Conditions.Fields)
                        {
                            if(c.FieldIndex == f.Index)
                            {
                                c.FieldIndex = ix;
                            }
                        }
                    }
                }
            }            
            fields.Clear();
            newFields.ForEach(x=>fields.Add(x));
        }

        private List<List<string>> GetTableFields()
        {
            var tableFields = new List<List<string>>();
            foreach(var tbl in _pivotTables)
            {
                var l = new List<string>();
                tableFields.Add(l);
                foreach(var field in tbl.Fields.OrderBy(x=>x.Index))
                {
                    l.Add(field.Name.ToLower());
                }
            }
            return tableFields;
        }

        private void RefreshPivotTableItems()
        {
            foreach(var pt in _pivotTables)
            {
                foreach(var fld in pt.Fields)
                {
                    fld.Items.Refresh();
                }
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
        internal void InitNew(ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, string xml)
        {
            var pck = pivotTable.WorkSheet._package.ZipPackage;

            CacheDefinitionXml = new XmlDocument();
            var sourceWorksheet = pivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheetName];
            if (xml == null)
            {
                LoadXmlSafe(CacheDefinitionXml, GetStartXml(sourceWorksheet, sourceAddress), Encoding.UTF8);
                TopNode = CacheDefinitionXml.DocumentElement;
            }
            else
            {
                CacheDefinitionXml = new XmlDocument();
                CacheDefinitionXml.LoadXml(xml);
                TopNode = CacheDefinitionXml.DocumentElement;
                
                string sourceName = SourceRange.GetName();
                if (string.IsNullOrEmpty(sourceName))
                {
                    SetXmlNodeString(_sourceWorksheetPath, sourceAddress.WorkSheetName);
                    SetXmlNodeString(_sourceAddressPath, sourceAddress.Address);
                }
                else
                {
                    SetXmlNodeString(_sourceNamePath, sourceName);
                }
            }

            CacheId = _wb.GetNewPivotCacheId();

            var c = CacheId;
            CacheDefinitionUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref c);
            Part = pck.CreatePart(CacheDefinitionUri, ContentTypes.contentTypePivotCacheDefinition);

            AddRecordsXml();
            LoadFields();
            CacheDefinitionXml.Save(Part.GetStream());
            _pivotTables.Add(pivotTable);
        }

        internal void ResetRecordXml(ZipPackage pck)
        {
            if (CacheRecordUri == null) return;

            var cacheRecord = new XmlDocument();
            cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");            ZipPackagePart recPart;

            if (pck.PartExists(CacheRecordUri))
            {
                recPart = pck.GetPart(CacheRecordUri);
            }
            else
            {
                recPart = pck.CreatePart(CacheRecordUri, ContentTypes.contentTypePivotCacheRecords); 
            }
            cacheRecord.Save(recPart.GetStream(FileMode.Create, FileAccess.Write));
        }
        private string GetStartXml(ExcelWorksheet sourceWorksheet, ExcelRangeBase sourceRange)
        {
            string xml = "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"6\" refreshedVersion=\"6\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

            xml += "<cacheSource type=\"worksheet\">";
            var sourceName = sourceRange.GetName();
            if (string.IsNullOrEmpty(sourceName))
            {
                xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceRange.Address, sourceRange.WorkSheetName);
            }
            else
            {
                xml += string.Format("<worksheetSource name=\"{0}\" /> ", sourceName);
            }
            xml += "</cacheSource>";
            xml += string.Format("<cacheFields count=\"{0}\">", sourceRange._toCol - sourceRange._fromCol + 1);
            for (int col = sourceRange._fromCol; col <= sourceRange._toCol; col++)
            {
                var name = sourceWorksheet?.GetValueInner(sourceRange._fromRow, col);
                if (name == null || name.ToString() == "")
                {
                    xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceRange._fromCol + 1);
                }
                else
                {
                    xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", SecurityElement.Escape(name.ToString()));
                }
                xml += "<sharedItems containsBlank=\"1\" /> ";
                xml += "</cacheField>";
            }
            xml += "</cacheFields>";
            xml += $"<extLst><ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{ExtLstUris.PivotCacheDefinitionUri}\"><x14:pivotCacheDefinition pivotCacheId=\"0\"/></ext></extLst>";
            xml += "</pivotCacheDefinition>";

            return xml;
        }
        internal void SetSourceName(string name)
        {
            DeleteNode(_sourceAddressPath); //Remove any address if previously set.
            SetXmlNodeString(_sourceNamePath, name);
        }
        internal void SetSourceAddress(string address)
        {
            DeleteNode(_sourceNamePath); //Remove any name or table if previously set.
            SetXmlNodeString(_sourceAddressPath, address);
        }
        int _cacheId = int.MinValue;
        internal int CacheId
        {
            get
            {
                if (_cacheId < 0)
                {
                    _cacheId = GetXmlNodeInt("d:extLst/d:ext/x14:pivotCacheDefinition/@pivotCacheId");
                    if (_cacheId < 0)
                    {
                        _cacheId = _wb.GetPivotCacheId(CacheDefinitionUri);
                        var node = GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");
                        node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{_cacheId}\"/>";
                    }
                }
                return _cacheId;
            }
            set
            {
                var node = GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");
                if(node.InnerXml=="")
                {
                    node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{_cacheId}\"/>";
                }
                else
                {
                    SetXmlNodeInt("d:extLst/d:ext/x14:pivotCacheDefinition/@pivotCacheId", value);
                }
            }
        }

        internal bool RefreshOnLoad 
        {
            get
            {
                return GetXmlNodeBool("@refreshOnLoad");
            }
            set
            {
                SetXmlNodeBool("@refreshOnLoad", value);
            }
        }

        public bool SaveData 
        { 
            get
            {
                return GetXmlNodeBool("@saveData", true);
            }
            set
            {
                if (SaveData == value) return;
                SetXmlNodeBool("@saveData", value);
                if (value)
                {
                    AddRecordsXml();
                }
                else
                {
                    RemoveRecordsXml();
                }
                SetXmlNodeBool("@saveData", value);
            }
        }

        private void RemoveRecordsXml()
        {
            RecordRelationshipId = null;
            _wb._package.ZipPackage.DeletePart(CacheRecordUri);
            CacheRecordUri = null;
            RecordRelationship = null;
        }

        private void AddRecordsXml()
        {
            int c = CacheId;
            //CacheRecord. Create an empty one.
            CacheRecordUri = GetNewUri(_wb._package.ZipPackage, "/xl/pivotCache/pivotCacheRecords{0}.xml", ref c);
            ResetRecordXml(_wb._package.ZipPackage);

            RecordRelationship = Part.CreateRelationship(UriHelper.ResolvePartUri(CacheDefinitionUri, CacheRecordUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
            RecordRelationshipId = RecordRelationship.Id;
        }

        internal void Delete()
        {
            _wb.RemovePivotTableCache(CacheId);
            Part.Package.DeletePart(CacheDefinitionUri);
            if (CacheRecordUri != null)
            {
                Part.Package.DeletePart(CacheRecordUri);
            }
        }
        internal ExcelPivotTableCacheField AddDateGroupField(ExcelPivotTableField field, eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int interval)
        {
            ExcelPivotTableCacheField cacheField = CreateField(groupBy.ToString(), field.Index, false);
            cacheField.SetDateGroup(field, groupBy, startDate, endDate, interval);

            Fields.Add(cacheField);
            return cacheField;
        }
        internal ExcelPivotTableCacheField AddFormula(string name, string formula)
        {
            ExcelPivotTableCacheField cacheField = CreateField(name, _fields.Count, false);
            cacheField.Formula = formula;
            Fields.Add(cacheField);
            return cacheField;
        }

        private ExcelPivotTableCacheField CreateField(string name, int index, bool databaseField=true)
        {
            //Add Cache definition field.
            var cacheTopNode = CacheDefinitionXml.SelectSingleNode("//d:cacheFields", NameSpaceManager);
            var cacheFieldNode = CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);
            
            cacheFieldNode.SetAttribute("name", name);
            if (databaseField == false)
            {
                cacheFieldNode.SetAttribute("databaseField", "0");
            }
            
            cacheTopNode.AppendChild(cacheFieldNode);
            
            return new ExcelPivotTableCacheField(NameSpaceManager, cacheFieldNode, this, index);
        }

    }
}
