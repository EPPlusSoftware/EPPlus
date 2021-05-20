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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    public class ExcelExternalWorkbook : ExcelExternalLink
    {
        Dictionary<string, int> _sheetNames = new Dictionary<string, int>();
        Dictionary<int, CellStore<object>> _sheetValues = new Dictionary<int, CellStore<object>>();
        Dictionary<int, CellStore<int>> _sheetMetaData = new Dictionary<int, CellStore<int>>();
        Dictionary<int, ExcelExternalNamedItemCollection<ExcelExternalDefinedName>> _definedNamesValues = new Dictionary<int, ExcelExternalNamedItemCollection<ExcelExternalDefinedName>>();
        HashSet<int> _sheetRefresh = new HashSet<int>();
        internal ExcelExternalWorkbook(ExcelWorkbook wb, ExcelPackage p) : base(wb)
        {
            SetPackage(p);
        }
        internal ExcelExternalWorkbook(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement)  : base(wb, reader, part, workbookElement)
        {
            var rId = reader.GetAttribute("id", ExcelPackage.schemaRelationships);
            Relation = part.GetRelationship(rId);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "sheetNames":
                            ReadSheetNames(reader);
                            break;
                        case "definedNames":
                            ReadDefinedNames(reader);
                            break;
                        case "sheetDataSet":
                            ReadSheetDataSet(reader, wb);
                            break;
                    }
                }
                else if(reader.NodeType==XmlNodeType.EndElement)
                {
                    if(reader.Name=="externalBook")
                    {
                        reader.Close();
                        break;
                    }
                }
            }

            CachedWorksheets = new ExcelExternalNamedItemCollection<ExcelExternalWorksheet>();
            CachedNames = GetNames(-1);
            foreach (var sheetName in _sheetNames.Keys)
            {
                var sheetId = _sheetNames[sheetName];
                CachedWorksheets.Add(new ExcelExternalWorksheet(
                       _sheetValues[sheetId], 
                       _sheetMetaData[sheetId],
                       _definedNamesValues[sheetId]) 
                { 
                    SheetId  = sheetId, 
                    Name =sheetName, 
                    RefreshError=_sheetRefresh.Contains(sheetId)
                });
            }
        }
        public override eExternalLinkType ExternalLinkType
        {
            get
            {
                return eExternalLinkType.ExternalWorkbook;
            }
        }

        private ExcelExternalNamedItemCollection<ExcelExternalDefinedName> GetNames(int ix)
        {
            if(_definedNamesValues.ContainsKey(ix))
            {
                return _definedNamesValues[ix];
            }
            else
            {
                return new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>();
            }
        }
        private void ReadSheetDataSet(XmlTextReader reader, ExcelWorkbook wb)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "sheetDataSet")
                {
                    break;
                }
                else if(reader.NodeType == XmlNodeType.Element && reader.Name == "sheetData")
                {
                    ReadSheetData(reader, wb);
                }
            }
        }
        private void ReadSheetData(XmlTextReader reader, ExcelWorkbook wb)
        {
            var sheetId = int.Parse(reader.GetAttribute("sheetId"));
            if(reader.GetAttribute("refreshError")=="1" && !_sheetRefresh.Contains(sheetId))
            {
                _sheetRefresh.Add(sheetId);
            }
            CellStore<object> cellStoreValues;
            CellStore<int> cellStoreMetaData;
            cellStoreValues = _sheetValues[sheetId];
            cellStoreMetaData = _sheetMetaData[sheetId];

            int row=0, col=0;
            string type="";
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "sheetData")
                {
                    break;
                }
                else if(reader.NodeType==XmlNodeType.Element)
                {
                    switch(reader.Name)
                    {
                        case "row":
                            row = int.Parse(reader.GetAttribute("r"));
                            break;
                        case "cell":
                            ExcelCellBase.GetRowCol(reader.GetAttribute("r"), out row, out col, false);
                            type = reader.GetAttribute("t");
                            var vm = reader.GetAttribute("vm");
                            if(!string.IsNullOrEmpty(vm))
                            {
                                cellStoreMetaData.SetValue(row, col, int.Parse(vm));
                            }
                            break;
                        case "v":
                            var v = ConvertUtil.GetValueFromType(reader, type, 0, wb);
                            cellStoreValues.SetValue(row, col, v);
                            break;
                    }
                }
            }
        }
        private void ReadDefinedNames(XmlTextReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "definedNames")
                {
                    break;
                }
                else if (reader.NodeType == XmlNodeType.Element && reader.Name == "definedName")
                {
                    int sheetId;
                    var sheetIdAttr = reader.GetAttribute("sheetId");
                    if (string.IsNullOrEmpty(sheetIdAttr))
                    {
                        sheetId = -1; // -1 represents the workbook level.
                    }
                    else
                    {
                        sheetId = int.Parse(sheetIdAttr);
                    }
                    
                    ExcelExternalNamedItemCollection<ExcelExternalDefinedName> names = _definedNamesValues[sheetId];

                    var name = reader.GetAttribute("name");
                    names.Add(new ExcelExternalDefinedName() { Name = reader.GetAttribute("name"), RefersTo = reader.GetAttribute("refersTo"), SheetId = sheetId });
                }
            }
        }
        private void ReadSheetNames(XmlTextReader reader)
        {
            var ix = 0;
            _definedNamesValues.Add(-1, new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>());
            while (reader.Read())
            {
                if(reader.NodeType==XmlNodeType.EndElement && reader.Name== "sheetNames")
                {
                    break;
                }
                else if(reader.NodeType==XmlNodeType.Element && reader.Name== "sheetName")
                {
                    _sheetValues.Add(ix, new CellStore<object>());
                    _sheetMetaData.Add(ix, new CellStore<int>());
                    _definedNamesValues.Add(ix, new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>());
                    _sheetNames.Add(reader.GetAttribute("val"), ix++);                    

                }
            }
        }
        /// <summary>
        /// An Uri to the external reference
        /// </summary>
        public Uri ExternalReferenceUri
        {
            get
            {
                return Relation.TargetUri;
            }
        }
        FileInfo _file=null;
        /// <summary>
        /// If the external reference is a file in the filesystem
        /// </summary>
        public FileInfo File
        {
            get
            {
                if(_file==null)
                {
                    var filePath = Relation?.TargetUri?.OriginalString;
                    if (string.IsNullOrEmpty(filePath) && HasWebProtocol(filePath)) return null;
                    if (filePath.StartsWith("file:///")) filePath = filePath.Substring(8);
                    try
                    {
                        
                        if(_wb._package.File!=null)
                        {
                            if (string.IsNullOrEmpty(Path.GetDirectoryName(filePath)) || Path.IsPathRooted(filePath) == false)
                            {
                                filePath = _wb._package.File.DirectoryName + "\\" + filePath;
                            }
                            else
                            {
                                if(Path.IsPathRooted(filePath) == true && filePath[0]==Path.DirectorySeparatorChar)
                                {
                                    filePath = _wb._package.File.Directory.Root.Name + filePath;
                                }
                            }
                        }                        
                        _file = new FileInfo(filePath);
                        if(!_file.Exists && _wb.ExternalReferences.Directories.Count>0)
                        {
                            SetDirectoryIfExists();
                        }
                    }
                    catch
                    {
                        return null;
                    }
                }
                return _file;
            }
            set
            {
                _file = value;
            }
        }

        private void SetDirectoryIfExists()
        {
            foreach(var d in _wb.ExternalReferences.Directories)
            {
                var file = d.FullName;
                if (file.EndsWith(Path.DirectorySeparatorChar.ToString()) == false)
                {
                    file += Path.DirectorySeparatorChar;
                }
                file += _file.Name;
                if (System.IO.File.Exists(file))
                {
                    _file = new FileInfo(file);
                    return;
                }
            }
        }

        ExcelPackage _package =null;
        public ExcelPackage Package
        {
            get
            {
                return _package;
            }
        }
        /// <summary>
        /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
        /// </summary>
        /// <returns>True if the load succeeded, otherwise false</returns>
        public bool Load()
        {
            if(File != null && File.Exists)
            {
                _package = new ExcelPackage(File);
                _package._loadedPackage = _wb._package;
                return true;
            }

            return false;
        }
        /// <summary>
        /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
        /// </summary>
        /// <returns>True if the load succeeded, otherwise false</returns>
        public bool Load(FileInfo packageFile)
        {
            if (packageFile.Exists)
            {
                SetPackage(new ExcelPackage(packageFile));
                return true;
            }
            return false;
        }
        /// <summary>
        /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
        /// </summary>
        /// <returns>True if the load succeeded, otherwise false</returns>
        public bool Load(ExcelPackage package)
        {
            if (package == null || package == _wb._package)
            {
                throw (new ArgumentException("The package can't be null or load itself."));
            }

            if (package.File == null)
            {
                throw (new ArgumentException("The package must have the File property set to be added as an external reference."));
            }

            SetPackage(package);

            return true;
        }

        private void SetPackage(ExcelPackage package)
        {
            _package = package;
            _package._loadedPackage = _wb._package;
        }

        /// <summary>
        /// Updates the external reference cache for the external workbook. To be used a <see cref="Package"/> must be loaded via the <see cref="Load"/> method.
        /// </summary>
        /// <returns>True if the update was successful otherwise false</returns>
        public bool UpdateCache()
        {            
            if (_package == null)
            {
                if(Load()==false)
                {
                    throw (new InvalidOperationException($"Can't update cache. The file {File?.FullName} does not exist."));
                }
            }

            var lexer = _wb.FormulaParser.Lexer;
            CachedWorksheets.Clear();
            CachedNames.Clear();
            foreach (var ws in _wb.Worksheets)
            {
                CachedWorksheets.Add(new ExcelExternalWorksheet() { Name = ws.Name, RefreshError = false });
            }
            foreach (var ws in _wb.Worksheets)
            {
                var formulas = new CellStoreEnumerator<object>(ws._formulas);
                foreach(var f in formulas)
                {
                    if(f is int sfIx)
                    {
                        var sf = ws._sharedFormulas[sfIx];
                        if(sf.Formula.Contains("["))
                        {
                            UpdateCacheForFormula(ws, sf.Formula, sf.Address);
                        }
                    }
                    else
                    {
                        var s = f.ToString();
                        if(s.Contains("["))
                        {
                            UpdateCacheForFormula(ws, s, "");
                        }
                    }
                }
            }
            return true;
        }

        private void UpdateCacheForFormula(ExcelWorksheet ws, string formula, string address)
        {
            var tokens = ws.Workbook.FormulaParser.Lexer.Tokenize(formula);

            foreach (var t in tokens)
            {
                if (t.TokenTypeIsSet(TokenType.ExcelAddress) || t.TokenTypeIsSet(TokenType.NameValue))
                {
                    if (ExcelCellBase.IsExternalAddress(t.Value))
                    {
                        if(t.TokenTypeIsSet(TokenType.ExcelAddress))
                        {
                            ExcelAddressBase a = new ExcelAddressBase(t.Value);
                            var ix = _wb.ExternalReferences.GetExternalReference(a._wb);
                            if (ix >= 0 && _wb.ExternalReferences[ix] == this)
                            {
                                UpdateCacheForAddress(a, address);
                            }
                        }
                        else
                        {
                            var name = t.Value.Substring(t.Value.IndexOf("]") + 1);
                            if (name.StartsWith("!")) name = name.Substring(1);
                            UpdateCacheForName(name);
                        }
                    }
                }
            }
        }

        private void UpdateCacheForName(string name)
        {
            int ix = 0;
            var wsName = ExcelAddressBase.GetWorksheetPart(name, "", ref ix);
            if (!string.IsNullOrEmpty(wsName))
            {
                name = name.Substring(ix);
            }

            ExcelNamedRange namedRange;
            if (string.IsNullOrEmpty(wsName))
            {
                namedRange = _package.Workbook.Names.ContainsKey(name) ? _package.Workbook.Names[name] : null;
            }
            else
            {
                var ws = _package.Workbook.Worksheets[wsName];
                if (ws == null)
                {
                    namedRange = null;
                }
                else
                {
                    namedRange = ws.Names.ContainsKey(name) ? ws.Names[name] : null;
                }
            }
            ExcelAddressBase referensTo;
            if(namedRange._fromRow>0)
            {
                referensTo = new ExcelAddressBase(namedRange.LocalAddress);
            }
            else
            {
                referensTo = new ExcelAddressBase("#REF!");
            }

            if(namedRange.LocalSheetId < 0)
            {
                if (!CachedNames.ContainsKey(name))
                {
                    CachedNames.Add(new ExcelExternalDefinedName() { Name = name, RefersTo = referensTo.Address, SheetId=-1 });
                    UpdateCacheForAddress(referensTo, "");
                }
            }
            else
            {
                var cws = CachedWorksheets[namedRange.LocalSheet.Name];
                if(cws != null)
                {
                    if (!CachedNames.ContainsKey(name))
                    {
                        cws.CachedNames.Add(new ExcelExternalDefinedName() { Name = name, RefersTo = referensTo.Address, SheetId = namedRange.LocalSheetId });
                        UpdateCacheForAddress(referensTo, "");
                    }
                }
            }
        }
        private void UpdateCacheForAddress(ExcelAddressBase formulaAddress, string sfAddress)
        {
            if (formulaAddress._fromRow < 0 || formulaAddress._fromCol < 0) return;
            if (string.IsNullOrEmpty(sfAddress) == false)
            {
                var a = new ExcelAddress(sfAddress);
                if (formulaAddress._toColFixed == false)
                {
                    formulaAddress._toCol += a.Columns - 1;
                    formulaAddress._toRow += a.Rows - 1;
                }
            }

            if (!string.IsNullOrEmpty(formulaAddress.WorkSheetName))
            {
                var ws = _package.Workbook.Worksheets[formulaAddress.WorkSheetName];
                if (ws == null)
                {
                    if (!CachedWorksheets.ContainsKey(formulaAddress.WorkSheetName))
                    {
                        CachedWorksheets.Add(new ExcelExternalWorksheet() { Name = ws.Name, RefreshError = true });
                    }
                }
                else
                {
                    var cws = CachedWorksheets[formulaAddress.WorkSheetName];
                    var cse = new CellStoreEnumerator<ExcelValue>(ws._values, formulaAddress._fromRow, formulaAddress._fromCol, formulaAddress._toRow, formulaAddress._toCol);
                    foreach(var v in cse)
                    {
                        cws.CellValues._values.SetValue(cse.Row, cse.Column, v._value);
                    }
                }
            }            
        }

        public override string ToString()
        {
            if (Relation?.TargetUri != null)
            {
                return ExternalLinkType.ToString() + "(" + Relation.TargetUri.ToString() + ")";
            }
            else
            {
                return base.ToString();
            }
        }
        internal ZipPackageRelationship Relation
        {
            get;
            set;
        }

        public ExcelExternalNamedItemCollection<ExcelExternalDefinedName> CachedNames
        {
            get;
        }
        public ExcelExternalNamedItemCollection<ExcelExternalWorksheet> CachedWorksheets
        {
            get;
        }

        internal override void Save(StreamWriter sw)
        {
            sw.Write($"<externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{Relation.Id}\">");
            sw.Write("<sheetNames>");
            foreach(var sheet in _sheetNames.OrderBy(x=>x.Value))
            {
                sw.Write($"<sheetName val=\"{ConvertUtil.ExcelEscapeString(sheet.Key)}\"/>");
            }
            sw.Write("</sheetNames><definedNames>");
            foreach (var sheet in _definedNamesValues.Keys)
            {
                foreach (ExcelExternalDefinedName name in _definedNamesValues[sheet])
                {
                    if(name.SheetId<0)
                    {
                        sw.Write($"<definedName name=\"{ConvertUtil.ExcelEscapeString(name.Name)}\" refersTo=\"{name.RefersTo}\" />");
                    }
                    else
                    {
                        sw.Write($"<definedName name=\"{ConvertUtil.ExcelEscapeString(name.Name)}\" refersTo=\"{name.RefersTo}\" sheetId=\"{name.SheetId:N0}\"/>");
                    }
                }
            }
            sw.Write("</definedNames><sheetDataSet>");
            foreach (var sheetId in _sheetValues.Keys)
            {
                sw.Write($"<sheetData sheetId=\"{sheetId}\"{(_sheetRefresh.Contains(sheetId) ? " refreshError=\"1\"" : "")}>");
                var cellEnum = new CellStoreEnumerator<object>(_sheetValues[sheetId]);
                var mdStore = _sheetMetaData[sheetId];
                var r = -1;
                while(cellEnum.Next())
                {
                    if(r!=cellEnum.Row)
                    {
                        if(r!=-1)
                        {
                            sw.Write("</row>");
                        }
                        sw.Write($"<row r=\"{cellEnum.Row}\">");                        
                    }
                    int md=-1;
                    if(mdStore.Exists(cellEnum.Row, cellEnum.Column, ref md))
                    {
                        sw.Write($"<cell r=\"{ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column)}\" md=\"{md}\"{ConvertUtil.GetCellType(cellEnum.Value, true)}><v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(cellEnum.Value, _wb.Date1904))}</v></cell>");
                    }
                    else
                    {
                        sw.Write($"<cell r=\"{ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column)}\"{ConvertUtil.GetCellType(cellEnum.Value, true)}><v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(cellEnum.Value, _wb.Date1904))}</v></cell>");
                    }
                    r = cellEnum.Row;
                }
                if (r != -1)
                {
                    sw.Write("</row>");
                }
                sw.Write("</sheetData>");
            }
            sw.Write("</sheetDataSet></externalBook>");            
        }
        
    }
}
