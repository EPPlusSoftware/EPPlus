using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.ExcelXMLWriter
{
    internal class ExtLstHelper
    {
        string initalXml;
        List<string> listOfExts = new List<string>();
        Dictionary<string, int> uriToIndex = new Dictionary<string, int>();

        public ExtLstHelper(string xml)
        {
            initalXml = xml;
            ParseIntialXmlToList(xml);
        }

        private void ParseIntialXmlToList(string xml)
        {
            int start = 0, end = 0;
            GetBlockPos(xml, "extLst", ref start, ref end);

            //If the node isn't just a placeholder
            if (end - start > 10)
            {
                int contentStart = start + "<ExtLst>".Length;
                string extNodesOnly = xml.Substring(contentStart, end - contentStart - "</ExtLst>".Length);

                string[] strLst = { "</ext>" };
                listOfExts = extNodesOnly.Split(strLst, StringSplitOptions.RemoveEmptyEntries).ToList();

                for (int i = 0; i < listOfExts.Count; i++)
                {
                    int startOfUri = listOfExts[i].LastIndexOf("{");
                    int endOfUri = listOfExts[i].LastIndexOf("}") + 1;

                    string uri = listOfExts[i].Substring(startOfUri, endOfUri - startOfUri);

                    uriToIndex.Add(uri, i);
                    listOfExts[i] += "</ext>";
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
            int indexOfNode = -1;
            if (uriOfNodeBefore != "")
            {
                indexOfNode = uriToIndex[uriOfNodeBefore];
            }

            List<string> keys = new List<string>(uriToIndex.Keys);

            if (indexOfNode == -1)
            {
                listOfExts.Insert(0, content);
                foreach (string key in keys)
                {
                    uriToIndex[key] += 1;
                }
                uriToIndex.Add(uri, 0);
            }
            else
            {
                if (indexOfNode + 1 > listOfExts.Count)
                {
                    listOfExts.Add(content);
                }
                else
                {
                    listOfExts.Insert(indexOfNode + 1, content);
                    foreach (string key in keys)
                    {
                        if (indexOfNode + 1 >= uriToIndex[key])
                        {
                            uriToIndex[key] += 1;
                        }
                    }
                }
                uriToIndex.Add(uri, indexOfNode + 1);
            }
        }

        internal string GetWholeExtLst()
        {
            string extLstString = "<extLst>";

            for (int i = 0; i < listOfExts.Count; i++)
            {
                extLstString += listOfExts[i];
            }

            extLstString += "</extLst>";
            return extLstString;
        }


        private void GetBlockPos(string xml, string tag, ref int start, ref int end)
        {
            Match startmMatch, endMatch;
            startmMatch = Regex.Match(xml.Substring(start), string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

            if (!startmMatch.Success) //Not found
            {
                start = -1;
                end = -1;
                return;
            }
            var startPos = startmMatch.Index + start;
            if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
            {
                end = startPos + startmMatch.Length;
            }
            else
            {
                endMatch = Regex.Match(xml.Substring(start), string.Format("(</[^>]*{0}[^>]*>)", tag));
                if (endMatch.Success)
                {
                    end = endMatch.Index + endMatch.Length + start;
                }
            }
            start = startPos;
        }

    }
}
