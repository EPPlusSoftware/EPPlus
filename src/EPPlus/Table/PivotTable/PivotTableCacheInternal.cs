using EPPlusTest.Table.PivotTable;
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
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting;
using System.Xml.XPath;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

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
            if (ExtLstCacheId <= 0)   //Check if the is set via exLst (used by for example slicers), otherwise set it to the cacheId
            {
                ExtLstCacheId = cacheId;
            }

            ZipPackageRelationship rel = Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheRecords").FirstOrDefault();
            if (rel != null)
            {
                CacheRecordUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            }

            _wb.SetNewPivotCacheId(cacheId);
        }
        internal const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";
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
        internal string SourceWorksheetName
        {
            get
            {
                return GetXmlNodeString(_sourceWorksheetPath);
            }
            set
            {
                SetXmlNodeString(_sourceWorksheetPath, value);
            }
        }

        internal const string _sourceRIdPath = "d:cacheSource/d:worksheetSource/@r:id";

        internal string SourceRId
        {
            get
            {
                return GetXmlNodeString(_sourceRIdPath);
            }
            set
            {
                SetXmlNodeString(_sourceRIdPath, value);
            }
        }

        internal ExcelRangeBase SourceRange 
        { 
            get
            {
                ExcelRangeBase sourceRange=null;
                if (CacheSource == eSourceType.Worksheet)
                {
                    if (SourceRId == null) //External workbook
                    {
                        return null;
                    }
                    else
                    {
                        var ws = _wb.Worksheets[SourceWorksheetName];
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
        /// The package internal URI to the pivot table cache definition Xml Document.
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
        internal PivotTableCacheRecords Records { get; private set; }
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
            var cacheNodes = CacheDefinitionXml.DocumentElement.SelectNodes("d:cacheFields/d:cacheField", NameSpaceManager);
            foreach (XmlNode node in cacheNodes)
            {
                _fields.Add(new ExcelPivotTableCacheField(NameSpaceManager, node, this, index++));
            }
            if(SaveData)
            {
                Records = new PivotTableCacheRecords(this);
            }
            else
            {
                Records = null;
            }
        }

        internal void RefreshFields(bool checkSourceValid)
        {
            if(checkSourceValid && IsSourceValid()==false) //If the source is not valid on save, skip refresh.
            {
                return;
            }
            UpdatePageFieldValues();
            var fields = new List<ExcelPivotTableCacheField>();
            var r = SourceRange;
            bool cacheUpdated=false;
            var  movedFields = new List<int>();
            var fieldsNode = GetNode("d:cacheFields");
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
                    if (_fields==null || ix >= _fields?.Count || _fields[ix].Name != name)
                    {
                        if (string.IsNullOrEmpty(name))
                        {
                            throw new InvalidOperationException($"Pivot Cache with id {CacheId} is invalid . Contains reference to a column with an empty header");
                        }
                        var fi = _fields.FindIndex(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                        if (fi<0)
                        {
                            field = CreateField(name, -1, true, true, ix == 0 ? null : fields[ix - 1].TopNode);
                            movedFields.Add(-1);
                            field.TopNode.InnerXml = "<sharedItems/>";
                        }
                        else
                        {
                            var x = 2;
                            while(movedFields.Contains(fi))
                            {
                                var dupName = name + x.ToString();
                                fi = _fields.FindIndex(x => x.Name.Equals(dupName, StringComparison.OrdinalIgnoreCase));
                                x++;
                            }
                            if(fi<0)
                            {
                                field = CreateField(name, -1, true, true, ix == 0 ? null : fields[ix - 1].TopNode);
                                field.TopNode.InnerXml = "<sharedItems/>";
                                movedFields.Add(-1);
                            }
                            else
                            {
                                field = _fields[fi];
                                movedFields.Add(fi);
                            }
                        }
                        field.SharedItems.Clear();
                        if (field._cacheLookup != null) field._cacheLookup.Clear();
                        if (cacheUpdated == false && string.IsNullOrEmpty(name) == false && !field.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase)) cacheUpdated = true;
                        cacheUpdated = true;
                    }
                    else
                    {
                        field = _fields[ix];
                        movedFields.Add(ix);
                        field.SharedItems.Clear();
                        if(field._cacheLookup!=null) field._cacheLookup.Clear();
                        if (cacheUpdated == false && string.IsNullOrEmpty(name)==false && !field.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase)) cacheUpdated=true;
                    }

                    var shNode = field.TopNode.SelectSingleNode("d:sharedItems", NameSpaceManager);
                    if (shNode.HasChildNodes)
                    {
                        shNode.RemoveAll();
                    }

                    if (!string.IsNullOrEmpty(name) && !field.Name.StartsWith(name)) field.Name = name;

                    if (cacheUpdated)
                    {
                        fieldsNode.RemoveChild(field.TopNode);
                        if (fields.Count == 0)
                        {
                            fieldsNode.PrependChild(field.TopNode);
                        }
                        else
                        {
                            fieldsNode.InsertAfter(field.TopNode, fields[fields.Count - 1].TopNode);
                        }
                    }

                    fields.Add(field);
                }
            }
            //Add non-database fields in the end.
            var i = _fields.Count - 1;
            var pos = fields.Count;
            while (i >= 0 && i < _fields.Count && _fields[i].DatabaseField == false)
            {
                fields.Insert(pos, _fields[i--]);
            }

            if (cacheUpdated || i >= fields.Count)
            {
                UpdateAndRemoveFields(fields, movedFields);
            }
            else
            {
                _fields = fields;
            }

            RefreshPivotTableItems();
            if (Records == null) Records = new PivotTableCacheRecords(this);
            Records.CreateRecords();
         }

        private void UpdateAndRemoveFields(List<ExcelPivotTableCacheField> fields, List<int> movedFields)
        {
            //Remove any fields from the existing list. 
            for (var i = 0; i < _fields.Count; i++)
            {
                if (!fields.Any(x => x.Name.Equals(_fields[i].Name, StringComparison.InvariantCultureIgnoreCase)))
                {
                    foreach (var pt in _pivotTables)
                    {
                        var node = pt.Fields[i].TopNode;
                        node.ParentNode.RemoveChild(node);
                    }
                    _fields[i].TopNode.ParentNode.RemoveChild(_fields[i].TopNode);
                }
            }
            _fields = fields;

            //Create new field elements and Move field elements in the list.
            List<List<ExcelPivotTableField>> pivotTableFields = new List<List<ExcelPivotTableField>>();
            _pivotTables.ForEach(x => pivotTableFields.Add(new List<ExcelPivotTableField>()));
            for (var i = 0; i < movedFields.Count; i++)
            {
                var ptIx = 0;
                foreach (var pt in _pivotTables)
                {
                    var list = pivotTableFields[ptIx++];
                    var parentNode = pt.GetNode("d:pivotFields");
                    var mi = movedFields[i];
                    if (mi == -1)
                    {
                        var node = pt.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
                        if (i == 0)
                        {
                            parentNode.PrependChild(node);
                        }
                        else
                        {
                            parentNode.InsertAfter(node, list[i - 1].TopNode);
                        }
                        var fld = new ExcelPivotTableField(pt.NameSpaceManager, node, pt, i, i);
                        //pt.Fields._list.Insert(i, fld);
                        list.Add(fld);
                        fld.Cache = null;
                    }
                    else
                    {
                        var field = pt.Fields._list[mi];
                        if (field.Index != i)
                        {
                            field.Index = i;
                            field.BaseIndex = i;
                            field.Cache = null;                            
                            field.TopNode.ParentNode.RemoveChild(field.TopNode);
                            var prevNode = i == 0 ? null : list[i - 1].TopNode;
                            if(prevNode==null)
                            {
                                parentNode.PrependChild(field.TopNode);
                            }
                            else
                            {
                                prevNode.ParentNode.InsertAfter(field.TopNode, prevNode);
                            }
                        }
                        list.Add(field);
                    }
                }
                fields[i].Index = i;                                    
            }
            for(int i=0;i<pivotTableFields.Count;i++)
            {
                _pivotTables[i].Fields._list = pivotTableFields[i];
            }

            UpdateFieldReferences(fields, movedFields);
        }

        private void UpdateFieldReferences(List<ExcelPivotTableCacheField> fields, List<int> movedFields)
        {
            var oldIndex = movedFields.Where(x => x >= 0).ToDictionary(x => movedFields.IndexOf(x), x => x);
            foreach (var pt in _pivotTables)
            {
                var rmDfFields = new List<ExcelPivotTableDataField>();
                //Update data field index
                foreach (var df in pt.DataFields)
                {
                    var ix = movedFields.IndexOf(df.Index);
                    if(ix<0)
                    {
                        rmDfFields.Add(df);
                    }
                    else if (df.Index != df.Field.Index)
                    {
                        df.Index = df.Field.Index;
                    }
                }

                rmDfFields.ForEach(df => { df.Field.IsDataField = false; pt.DataFields.Remove(df); });
                if(pt.DataFields.Count==0)
                {
                    pt.DeleteNode("d:dataFields");
                }

                //Update column field index
                foreach (var cf in pt.ColumnFields)
                {
                    UpdateRowColPathFieldXml(oldIndex, pt, cf, "d:colFields/d:field[@x={0}]", "x");
                }

                //Update row field index
                foreach (var rf in pt.RowFields)
                {
                    UpdateRowColPathFieldXml(oldIndex, pt, rf, "d:rowFields/d:field[@x={0}]", "x");
                }

                //Update page field index
                foreach (var pf in pt.PageFields)
                {
                    UpdateRowColPathFieldXml(oldIndex, pt, pf, "d:pageFields/d:pageField[@fld={0}]", "fld");
                }

                //Update styles
                var newIndex = movedFields.Where(x => x >= 0).ToDictionary(x => x, x => movedFields.IndexOf(x));
                foreach (ExcelPivotTableAreaStyle s in pt.Styles)
                {
                    if(s.FieldIndex.HasValue && s.FieldIndex.Value>=0)
                    {
                        if (newIndex.TryGetValue(s.FieldIndex.Value, out int newIx))
                        {
                            s.FieldIndex = newIx;
                        }
                        else
                        {
                            //Remove
                        }
                    }

                    foreach (var f in s.Conditions.Fields)
                    {
                        if (newIndex.TryGetValue(f.Field.Index, out int newIx))
                        {
                            f.FieldIndex = newIx;
                        }
                    }

                    s.Conditions.UpdateXml();
                }

                //Update conditional formatting
                foreach (var fs in pt.ConditionalFormattings)
                {
                    foreach (var a in fs.Areas)
                    {
                        if (a.FieldIndex.HasValue && a.FieldIndex.Value >= 0)
                        {
                            if (newIndex.TryGetValue(a.FieldIndex.Value, out int newIx))
                            {
                                a.FieldIndex = newIx;
                            }
                            else
                            {
                                //Remove
                            }
                        }
                        foreach (var f in a.Conditions.Fields)
                        {
                            if (newIndex.TryGetValue(f.Field.Index, out int newIx))
                            {
                                f.FieldIndex = newIx;
                            }
                        }

                        a.Conditions.UpdateXml();
                    }
                }
            }
        }

        private static void UpdateRowColPathFieldXml(Dictionary<int, int> oldIndex, ExcelPivotTable pt, ExcelPivotTableField ptf, string xpath, string attributeName)
        {
            if (ptf.Index >= 0)
            {
                if (oldIndex.TryGetValue(ptf.Index, out int oldIx))
                {
                    if (ptf.Index != oldIx)
                    {
                        var node = pt.GetNode(string.Format(xpath, oldIx)) as XmlElement;
                        if (node != null)
                        {
                            node.SetAttribute(attributeName, $"{ptf.Index}");
                        }
                    }
                }
                else
                {
                    var node = pt.GetNode(string.Format(xpath, ptf.Index)) as XmlElement;
                    if (node != null)
                    {
                        var parent = node.ParentNode;
                        parent.RemoveChild(node);
                        if (parent.ChildNodes.Count == 0)
                        {
                            parent.ParentNode.RemoveChild(parent);
                        }
                    }
                }
            }
        }

        private void UpdatePageFieldValues()
        {
            foreach(var pt in _pivotTables)
            {
                foreach(var pf in pt.PageFields)
                {
                    if (pf.PageFieldSettings.SelectedItem>=0 && pf.PageFieldSettings.SelectedItem < pf.Items.Count)
                    {
                        pf.PageFieldSettings.SelectedValue = pf.Items[pf.PageFieldSettings.SelectedItem].Value;
                    }
                }
            }
        }

        private void RefreshPivotTableItems()
        {
            foreach(var pt in _pivotTables)
            {
                if (pt.CacheDefinition.CacheSource == eSourceType.Worksheet)
                {
                    for(int i=0;i < pt.Fields.Count; i++)
                    {
                        var field = pt.Fields[i];
                        field.Items.Refresh();
                        if(field.IsPageField && field.PageFieldSettings.SelectedItem > -1)
                        {
                            field.PageFieldSettings.SelectedItem = field.Items.GetIndexByValue(field.PageFieldSettings.SelectedValue);
                        }
                    }
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
                    SourceWorksheetName = sourceAddress.WorkSheetName;
                    SetXmlNodeString(_sourceAddressPath, sourceAddress.Address);
                }
                else
                {
                    SetXmlNodeString(_sourceNamePath, sourceName);
                }
            }

            CacheId = ExtLstCacheId = _wb.GetNewPivotCacheId();

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
                var innerValue = sourceWorksheet?.GetValueInner(sourceRange._fromRow, col);
                string name = "";
                if (sourceWorksheet._flags.GetFlagValue(sourceRange._fromRow, col, CellFlags.RichText))
                {
                    name = sourceWorksheet.GetRichText(sourceRange._fromRow, col, sourceWorksheet.Cells[sourceRange._fromRow, col]).Text;
                }
                else
                {
                    name = innerValue.ToString();
                }

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

        /// <summary>
        /// This is the cache id from the workbook 
        /// </summary>
        internal int CacheId
        {
            get
            {
                if (_cacheId < 0)
                {
                    _cacheId = _wb.GetPivotCacheId(CacheDefinitionUri);
                }
                return _cacheId;
            }
            set
            {
                _cacheId = value;
            }
        }

        int _extLstCacheId = int.MinValue;
        /// <summary>
        /// This a second cache id used for newer items like slicers. EPPlus will set this id to the same as the cache id by default.
        /// </summary>
        internal int ExtLstCacheId
        {
            get
            {
                if (_extLstCacheId < 0)
                {
                    _extLstCacheId = GetXmlNodeInt("d:extLst/d:ext/x14:pivotCacheDefinition/@pivotCacheId");
                    if (_extLstCacheId < 0)
                    {
                        _extLstCacheId = CacheId;
                        var node = GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");
                        node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{_extLstCacheId}\"/>";
                    }
                }
                return _extLstCacheId;
            }
            set
            {
                var node = GetOrCreateExtLstSubNode(ExtLstUris.PivotCacheDefinitionUri, "x14");
                if (node.InnerXml == "")
                {
                    node.InnerXml = $"<x14:pivotCacheDefinition pivotCacheId=\"{_extLstCacheId}\"/>";
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

        public Uri SourceExternalReferenceUri 
        {
            get
            {
                if (string.IsNullOrEmpty(SourceRId) || Part.RelationshipExists(SourceRId) == false)
                {
                    return null;
                }
                else
                {
                    return Part.GetRelationship(SourceRId).TargetUri;
                }

            }
        }

        public bool IsSourceValid()
        {
            var r = SourceRange;
            for (int col = r._fromCol; col <= r._toCol; col++)
            {
                var ix = col - r._fromCol;
                var ws = r.Worksheet;
                var name = ws.GetValue(r._fromRow, col)?.ToString().Trim();
                if (string.IsNullOrEmpty(name))
                {
                    return false;
                }
            }
            return true;
        }

        private void RemoveRecordsXml()
        {
            RecordRelationshipId = null;
            _wb._package.ZipPackage.DeletePart(CacheRecordUri);
            CacheRecordUri = null;
            RecordRelationship = null;
        }

        internal void AddRecordsXml()
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
            cacheField.SetDateGroup(field, groupBy, startDate, endDate, interval, false);

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
        private ExcelPivotTableCacheField CreateField(string name, int index, bool databaseField=true, bool insertAfter=false, XmlNode prependingNode=null)
        {
            //Add Cache definition field.
            var cacheTopNode = CacheDefinitionXml.SelectSingleNode("//d:cacheFields", NameSpaceManager);
            var cacheFieldNode = CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);
            
            cacheFieldNode.SetAttribute("name", name);
            if (databaseField == false)
            {
                cacheFieldNode.SetAttribute("databaseField", "0");
            }
            if(insertAfter)
            {
                cacheTopNode.InsertAfter(cacheFieldNode, prependingNode);
            }
            else
            {
                cacheTopNode.AppendChild(cacheFieldNode);
            }

            return new ExcelPivotTableCacheField(NameSpaceManager, cacheFieldNode, this, index);
        }

        internal string GetSourceAddress()
        {

            if (string.IsNullOrEmpty(SourceRId))
            {
                if (string.IsNullOrEmpty(SourceName))
                {
                    var r=SourceRange?.FullAddress;
                    if(r==null)
                    {
                        if(string.IsNullOrEmpty(SourceWorksheetName))
                        {
                            return Ref;
                        }
                        else
                        {
                            return ExcelCellBase.GetQuotedWorksheetName(SourceWorksheetName) + "!" + Ref;
                        }                            
                    }
                    return r;
                }
                else
                {
                    return SourceName;
                }
            }
            else
            {
                var uri = SourceExternalReferenceUri;
                if (uri!=null)
                {
                    var refIx=_wb.ExternalLinks.GetExternalLink(uri.OriginalString);
                    if(refIx>=0)
                    {
                        return $"[{refIx}]{SourceWorksheetName}!{Ref}";
                    }
                    else
                    {
                        try
                        {
                            var fi = new FileInfo(uri.OriginalString);
                            return $"[{fi.Name}]{SourceWorksheetName}!{Ref}";
                        }
                        catch
                        {
                            return $"[{uri.OriginalString}]{SourceWorksheetName}!{Ref}";
                        }
                    }
                }
                return null;
            }
        }

		internal int GetMaxRow()
		{
            var range = SourceRange;

			var dimensionToRow = range.Worksheet?.Dimension?._toRow + 1 ?? range._fromRow + 1; //We add 1 to dimension to row so we get one row with null values.
			var toRow = range._toRow < dimensionToRow ? range._toRow : dimensionToRow;
			return toRow;
		}
	}
}
