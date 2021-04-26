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
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Core.ExternalReferences
{
    public class ExcelExternalReference 
    {
        Dictionary<string, int> _sheetNames = new Dictionary<string, int>();
        Dictionary<int, CellStore<object>> _sheetValues = new Dictionary<int, CellStore<object>>();
        Dictionary<int, CellStore<int>> _sheetMetaData = new Dictionary<int, CellStore<int>>();
        Dictionary<int, ExcelNamedItemCollection<ExcelExternalDefinedName>> _definedNamesValues = new Dictionary<int, ExcelNamedItemCollection<ExcelExternalDefinedName>>();
        internal ExcelWorkbook _wb;
        internal XmlElement WorkbookElement
        {
            get;
            set;
        }

        internal ZipPackagePart Part 
        {
            get;
            set;
        }
        internal ZipPackageRelationship Relation
        {
            get;
            set;
        }
        internal ExcelExternalReference(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement) 
        {
            if (reader.LocalName != "externalBook") return;
            var rId = reader.GetAttribute("id", ExcelPackage.schemaRelationships);
            _wb = wb;
            Part = part;
            WorkbookElement = workbookElement;
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

            CachedWorksheets = new ExcelNamedItemCollection<ExcelExternalWorksheet>();
            CachedNames = GetNames(-1);
            foreach (var sheetName in _sheetNames.Keys)
            {
                var sheetId = _sheetNames[sheetName];
                CachedWorksheets.Add(new ExcelExternalWorksheet(
                       _sheetValues[sheetId], 
                       _sheetMetaData[sheetId],
                       _definedNamesValues[sheetId]) { SheetId  = sheetId, Name =sheetName});
            }
        }
        private ExcelNamedItemCollection<ExcelExternalDefinedName> GetNames(int ix)
        {
            if(_definedNamesValues.ContainsKey(ix))
            {
                return _definedNamesValues[ix];
            }
            else
            {
                return new ExcelNamedItemCollection<ExcelExternalDefinedName>();
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
                    
                    ExcelNamedItemCollection<ExcelExternalDefinedName> names = _definedNamesValues[sheetId];

                    var name = reader.GetAttribute("name");
                    names.Add(new ExcelExternalDefinedName() { Name = reader.GetAttribute("name"), RefersTo = reader.GetAttribute("refersTo"), SheetId = sheetId });
                }
            }
        }
        private void ReadSheetNames(XmlTextReader reader)
        {
            var ix = 0;
            _definedNamesValues.Add(-1, new ExcelNamedItemCollection<ExcelExternalDefinedName>());
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
                    _definedNamesValues.Add(ix, new ExcelNamedItemCollection<ExcelExternalDefinedName>());
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
        ExcelPackage _package=null;
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
            var fi = new FileInfo(ExternalReferenceUri.LocalPath);
            if(fi.Exists)
            {
                _package = new ExcelPackage(fi);
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
                _package = new ExcelPackage(packageFile);
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
            if (package == null || package.File==null)
            {
                throw (new ArgumentException("The package must have the File property set to be added as an external reference."));
            }
            _package = package;

            return true;
        }

        public override string ToString()
        {
            if (Relation?.TargetUri != null)
            {
                return Relation.TargetUri.ToString();
            }
            else
            {
                return base.ToString();
            }
        }        
        public ExcelNamedItemCollection<ExcelExternalDefinedName> CachedNames
        {
            get;
        }
        public ExcelNamedItemCollection<ExcelExternalWorksheet> CachedWorksheets
        {
            get;
        }

        internal void Save()
        {
            var sw = new StreamWriter(Part.GetStream(FileMode.CreateNew));            
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{Relation.Id}\">");
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
                sw.Write($"<sheetData sheetId=\"{sheetId}\">");
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
            sw.Write("</sheetDataSet></externalBook></externalLink>");
            sw.Flush();
        }
        
    }
}
