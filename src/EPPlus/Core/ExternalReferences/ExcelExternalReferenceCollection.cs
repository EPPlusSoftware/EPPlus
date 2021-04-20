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
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Core.ExternalReferences
{
    public class ExcelExternalReferenceCollection : IEnumerable<ExcelExternalReference>
    {
        List<ExcelExternalReference> _list=new List<ExcelExternalReference>();
        ExcelWorkbook _wb;
        internal ExcelExternalReferenceCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            GetExternalReferences();
        }
        internal void AddInternal(ExcelExternalReference externalReference)
        {
            _list.Add(externalReference);
        }
        public IEnumerator<ExcelExternalReference> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public int Count { get { return _list.Count; } }
        public ExcelExternalReference this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        internal void GetExternalReferences()
        {
            XmlNodeList nl = _wb.WorkbookXml.SelectNodes("//d:externalReferences/d:externalReference", _wb.NameSpaceManager);
            if (nl != null)
            {
                foreach (XmlElement elem in nl)
                {
                    string rID = elem.GetAttribute("r:id");
                    var rel = _wb.Part.GetRelationship(rID);
                    var part = _wb._package.ZipPackage.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                    var xr = new XmlTextReader(part.GetStream());
                    while (xr.Read())
                    {
                        if (xr.NodeType == XmlNodeType.Element)
                        {
                            switch (xr.Name)
                            {
                                case "externalBook":
                                    AddInternal(new ExcelExternalReference(_wb, xr, part));
                                    break;
                                case "ddeLink":
                                case "oleLink":
                                case "extLst":
                                    break; //Unsupported
                                default:    //If we end up here the workbook is invalid.
                                    break;
                            }
                        }
                    }
                }
            }
        }
    }
}
