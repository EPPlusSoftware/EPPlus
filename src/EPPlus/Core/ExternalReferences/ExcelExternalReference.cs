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
using System.Xml;

namespace OfficeOpenXml.Core.ExternalReferences
{
    public class ExcelExternalReference 
    {
        Dictionary<string, int> _sheetNames = new Dictionary<string, int>();
        Dictionary<int, CellStore<object>> _sheetValues = new Dictionary<int, CellStore<object>>();
        Dictionary<int, CellStore<int>> _sheetMetaData = new Dictionary<int, CellStore<int>>();
        Dictionary<int, ExcelNamedItemCollection<ExcelExternalDefinedName>> _definedNamesValues = new Dictionary<int, ExcelNamedItemCollection<ExcelExternalDefinedName>>();

        internal ZipPackageRelationship _externalRelation;
        internal ExcelExternalReference(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part) 
        {
            if (reader.LocalName != "externalBook") return;
            var rId = reader.GetAttribute("id", ExcelPackage.schemaRelationships);
            _externalRelation = part.GetRelationship(rId);
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

            Worksheets = new ExcelNamedItemCollection<ExcelExternalWorksheet>();
            Names = GetNames(-1);
            foreach (var sheetName in _sheetNames.Keys)
            {
                var sheetId = _sheetNames[sheetName];
                Worksheets.Add(new ExcelExternalWorksheet(
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
                return _externalRelation.TargetUri;
            }
        }
        public override string ToString()
        {
            if (_externalRelation?.TargetUri != null)
            {
                return _externalRelation.TargetUri.ToString();
            }
            else
            {
                return base.ToString();
            }
        }        
        public ExcelNamedItemCollection<ExcelExternalDefinedName> Names
        {
            get;
        }
        public ExcelNamedItemCollection<ExcelExternalWorksheet> Worksheets
        {
            get;
        }
    }
}
