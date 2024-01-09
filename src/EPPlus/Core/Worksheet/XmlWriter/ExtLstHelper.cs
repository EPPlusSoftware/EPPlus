/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2023         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.XmlWriter;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.ExcelXMLWriter
{
    internal class ExtLstHelper
    {
        List<string> _listOfExts = new List<string>();
        internal int extCount { get { return _listOfExts.Count; } }

        Dictionary<string, int> _uriToIndex = new Dictionary<string, int>();

        public ExtLstHelper(string xml, ExcelWorksheet ws)
        {
            ParseIntialXmlToList(xml, ws);
        }

        private void ParseIntialXmlToList(string xml, ExcelWorksheet ws)
        {
            int start = 0, end = 0;
            GetBlock.Pos(xml, "extLst", ref start, ref end);

            bool isPlaceHolder = false;

            if (!xml.Substring(start + 1, end - start - 1).Contains("<"))
            {
                isPlaceHolder = true;
            }

            //If the node isn't just a placeholder
            if (!isPlaceHolder)
            {
                int contentStart = start + "<ExtLst>".Length;
                string extNodesOnly = xml.Substring(contentStart, end - contentStart - "</ExtLst>".Length);

                string[] strLst = { "</ext>" };
                _listOfExts = extNodesOnly.Split(strLst, StringSplitOptions.RemoveEmptyEntries).ToList();

                for (int i = 0; i < _listOfExts.Count; i++)
                {
                    int startOfUri = _listOfExts[i].LastIndexOf("{");
                    int endOfUri = _listOfExts[i].LastIndexOf("}") + 1;

                    string uri;

					if (startOfUri >= 0)
                    {
						uri = _listOfExts[i].Substring(startOfUri, endOfUri - startOfUri);
					}
                    else
                    {
						_uriToIndex.Add(i.ToString(), i);
					}

					_listOfExts[i] += "</ext>";
                }
            }
        }
        /// <summary>
        /// Inserts content after the uriNode
        /// Note that this is only intended to be done once per type of node and it will throw error
        /// if the same uri is attempted in two separate calls or if it's already been read in initally.
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="content"></param>
        /// If <param name="uriOfNodeBefore"> is blank sets content as the first ext</param>
        internal void InsertExt(string uri, string content, string uriOfNodeBefore)
        {
            if (string.IsNullOrEmpty(content))
            {
                return;
            }

            int indexOfNode = -1;
            if (uriOfNodeBefore != "")
            {
                indexOfNode = _uriToIndex[uriOfNodeBefore];
            }

            List<string> keys = new List<string>(_uriToIndex.Keys);

            if (indexOfNode == -1)
            {
                _listOfExts.Insert(0, content);
                foreach (string key in keys)
                {
                    _uriToIndex[key] += 1;
                }
                _uriToIndex.Add(uri, 0);
            }
            else
            {
                if (indexOfNode + 1 > _listOfExts.Count)
                {
                    _listOfExts.Add(content);
                }
                else
                {
                    _listOfExts.Insert(indexOfNode + 1, content);
                    foreach (string key in keys)
                    {
                        if (indexOfNode + 1 >= _uriToIndex[key])
                        {
                            _uriToIndex[key] += 1;
                        }
                    }
                }
                _uriToIndex.Add(uri, indexOfNode + 1);
            }
        }

        internal string GetWholeExtLst()
        {
            string extLstString = "<extLst>";

            for (int i = 0; i < _listOfExts.Count; i++)
            {
                extLstString += _listOfExts[i];
            }

            extLstString += "</extLst>";
            return extLstString;
        }
    }
}
