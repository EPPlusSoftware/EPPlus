/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sorting
{
    public class SortCondition : XmlHelper
    {
        public SortCondition(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }

        private string _descendingPath = "@descending";
        private string _refPath = "@ref";
        private string _customListPath = "@customList";

        public bool Descending
        {
            get
            {
                return GetXmlNodeBool(_descendingPath);
            }
            set
            {
                SetXmlNodeBool(_descendingPath, value);
            }
        }

        public string Ref 
        {
            get
            {
                return GetXmlNodeString(_refPath);
            }
            set
            {
                SetXmlNodeString(_refPath, value);
            }
        }

        public string[] CustomList
        {
            get
            {
                var list = GetXmlNodeString(_customListPath);
                if(!string.IsNullOrEmpty(list))
                {
                    return list.Split(',').Where(x => !string.IsNullOrEmpty(x)).Select(x => x.Trim()).ToArray();
                }
                return null;
            }
            set
            {
                if(value == null || value.Length == 0)
                {
                    SetXmlNodeString(_customListPath, string.Empty, true);
                }
                var val = new StringBuilder();
                for(var x = 0; x < value.Length; x++)
                {
                    val.Append(value[x]);
                    if(x < value.Length -1)
                    {
                        val.Append(",");
                    }
                }
                SetXmlNodeString(_customListPath, val.ToString());
            }
        }
    }
}
