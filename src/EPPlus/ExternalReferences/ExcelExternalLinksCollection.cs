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
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    public class ExcelExternalLinksCollection : IEnumerable<ExcelExternalLink>
    {
        List<ExcelExternalLink> _list=new List<ExcelExternalLink>();
        ExcelWorkbook _wb;
        internal ExcelExternalLinksCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            LoadExternalReferences();
        }
        internal void AddInternal(ExcelExternalLink externalLink)
        {
            _list.Add(externalLink);
        }
        public IEnumerator<ExcelExternalLink> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public int Count { get { return _list.Count; } }
        public ExcelExternalLink this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        /// <summary>
        /// Adds an external reference to another workbook. 
        /// </summary>
        /// <param name="file">The location of the external workbook. The external workbook must of type .xlsx, .xlsm or xlst</param>
        /// <returns>The <see cref="ExcelExternalWorkbook"/> object</returns>
        public ExcelExternalWorkbook AddExternalWorkbook(FileInfo file)
        {
            if(file == null || file.Exists==false)
            {
                throw (new FileNotFoundException("The file does not exist."));
            }
            var p = new ExcelPackage(file);
            var ewb = new ExcelExternalWorkbook(_wb, p);
            _list.Add(ewb);
            return ewb;
        }
        internal void LoadExternalReferences()
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
                                    AddInternal(new ExcelExternalWorkbook(_wb, xr, part, elem));
                                    break;
                                case "ddeLink":
                                    AddInternal(new ExcelExternalDdeLink(_wb, xr, part, elem));
                                    break;
                                case "oleLink":
                                    AddInternal(new ExcelExternalOleLink(_wb, xr, part, elem));
                                    break;
                                case "extLst":

                                    break; 
                                default:    
                                    break;
                            }
                        }
                    }
                    xr.Close();
                }
            }
        }
        /// <summary>
        /// Removes the external link at the zero-based index. If the external reference is an workbook any formula links are broken.
        /// </summary>
        /// <param name="index">The zero-based index</param>
        public void RemoveAt(int index)
        {
            if(index < 0 || index>=_list.Count)
            {
                throw (new ArgumentOutOfRangeException(nameof(index)));
            }
            Remove(_list[index]);
        }
        /// <summary>
        /// Removes the external link from the package.If the external reference is an workbook any formula links are broken.
        /// </summary>
        /// <param name="externalLink"></param>
        public void Remove(ExcelExternalLink externalLink)
        {
            var ix = _list.IndexOf(externalLink);
            
            _wb._package.ZipPackage.DeletePart(externalLink.Part.Uri);

            if(externalLink.ExternalLinkType==eExternalLinkType.ExternalWorkbook)
            {
                ExternalLinksHandler.BreakFormulaLinks(_wb, ix, true);
            }

            var extRefs = externalLink.WorkbookElement.ParentNode;
            extRefs?.RemoveChild(externalLink.WorkbookElement);
            if(extRefs?.ChildNodes.Count==0)
            {
                extRefs.ParentNode?.RemoveChild(extRefs);
            }
            _list.Remove(externalLink);
        }
        /// <summary>
        /// Clear all external links and break any formula links.
        /// </summary>
        public void Clear()
        {
            if (_list.Count == 0) return;
            var extRefs = _list[0].WorkbookElement.ParentNode;

            ExternalLinksHandler.BreakAllFormulaLinks(_wb);
            while (_list.Count>0)
            {
                _wb._package.ZipPackage.DeletePart(_list[0].Part.Uri);
                _list.RemoveAt(0);
            }

            extRefs?.ParentNode?.RemoveChild(extRefs);
        }
        /// <summary>
        /// A list of directories to look for the external files that cannot be found on the path of the uri.
        /// </summary>
        public List<DirectoryInfo> Directories
        {
            get;
        } = new List<DirectoryInfo>();
        /// <summary>
        /// Will load all external workbooks that can be accessed via the file system.
        /// External workbook referenced via other protocols must be loaded manually.
        /// </summary>
        /// <returns>Returns false if any workbook fails to loaded otherwise true. </returns>
        public bool LoadWorkbooks()
        {
            bool ret = true;
            foreach (var link in _list)
            {
                if(link.ExternalLinkType==eExternalLinkType.ExternalWorkbook)
                {
                    var externalWb = link.As.ExternalWorkbook;
                    if(externalWb.Package==null)
                    {
                        if(externalWb.Load() == false)
                        {
                            ret = false;
                        }
                    }
                }
            }
            return ret;
        }
        internal int GetExternalLink(string extRef)
        {
            if (string.IsNullOrEmpty(extRef)) return -1;
            if(extRef.Any(c=>char.IsDigit(c)==false))
            {
                if(ExcelExternalLink.HasWebProtocol(extRef))
                {
                    for (int ix = 0; ix < _list.Count; ix++)
                    {
                        if (_list[ix].ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                        {
                            if (extRef.Equals(_list[ix].As.ExternalWorkbook.ExternalLinkUri.OriginalString, StringComparison.OrdinalIgnoreCase))
                            {
                                return ix;
                            }
                        }
                    }
                    return -1;
                }
                if (extRef.StartsWith("file:///")) extRef = extRef.Substring(8);

                int ret = -1;
                try
                {
                    var fi = new FileInfo(extRef);
                    for (int ix = 0; ix < _list.Count; ix++)
                    {
                        if (_list[ix].ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                        {

                            var wb = _list[ix].As.ExternalWorkbook;
                            if (wb.File == null)
                            {
                                var fileName = wb.ExternalLinkUri?.OriginalString;
                                if (ExcelExternalLink.HasWebProtocol(fileName))
                                {
                                    if (fileName.Equals(extRef, StringComparison.OrdinalIgnoreCase))
                                    {
                                        return ix;
                                    }
                                    continue;
                                }
                            }

                            if (fi.Name.Equals(wb.File.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                ret = ix;
                            }
                        }
                    }
                }
                catch   //If the FileInfo is 
                {
                    return -1;
                }
                return ret;
            }
            else
            {
                var ix = int.Parse(extRef)-1;
                if(ix<_list.Count)
                {
                    return ix;
                }
            }
            return -1;
        }
        internal int GetIndex(ExcelExternalLink link)
        {
            return _list.IndexOf(link);
        }
        /// <summary>
        /// Updates the value cache for any external workbook in the collection. The link must be an workbook and of type xlsx, xlsm or xlst.
        /// </summary>
        /// <returns>True if all updates succeeded, otherwise false. Any errors can be found on the External links. <seealso cref="ExcelExternalLink.ErrorLog"/></returns>
        public bool UpdateCaches()
        {
            var ret = true;
            foreach(var er in _list)
            {
                if(er.ExternalLinkType==eExternalLinkType.ExternalWorkbook)
                {
                    if(er.As.ExternalWorkbook.UpdateCache()==false)
                    {
                        ret = false;
                    }
                }
            }
            return ret;
        }
    }
}
