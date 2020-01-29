/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Xml;
using System.Globalization;
using System;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A numeric source for a chart.
    /// </summary>
    public class ExcelChartNumericSource : XmlHelper
    {
        string _path;
        XmlElement _sourceElement=null;
        string _formatCodePath;
        internal ExcelChartNumericSource(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder) : base(nameSpaceManager, topNode)
        {
            _path = path;
            _formatCodePath = $"{_path}/c:numLit/c:formatCode";
            AddSchemaNodeOrder(schemaNodeOrder,new string[] { "formatCode", "ptCount", "pt" });
            SetSourceElement();
            if (_sourceElement != null)
            {
                switch (_sourceElement.LocalName)
                {
                    case "numLit":
                        _formatCode = GetXmlNodeString(_path + "/c:numLit/c:formatCode");
                        break;
                    case "numRef":
                        _formatCode = GetXmlNodeString(_path + "/c:numRef/c:numCache/c:formatCode");
                        break;
               }
            }
        }
        /// <summary>
        /// This can be an address, function or litterals.
        /// Litternals are formatted as a comma separated list surrounded by curly brackets, for example {1.0,2.0,3}. Please use a dot(.) as decimal sign.
        /// </summary>
        public string ValuesSource
        {
            get
            {
                if(_sourceElement==null)
                {
                    return "";
                }
                else if(_sourceElement.LocalName=="numLit")
                {
                    return GetNumLit();
                }
                else
                {
                    return GetXmlNodeString($"{_path}/c:numRef/c:f");
                }
            }
            set
            {
                if (_sourceElement != null) _sourceElement.ParentNode.RemoveChild(_sourceElement);

                value = value.Trim();
                if (value.StartsWith("=")) value = value.Substring(1);

                if (value.StartsWith("{"))
                {
                    if (!value.EndsWith("}")) throw new ArgumentException("ValueSource", "Invalid format:Litteral values must begin and end with a curly bracket");
                    CreateNumLit(value);
                }
                else
                {
                    SetXmlNodeString($"{_path}/c:numRef/c:f", value);
                }
                if (!string.IsNullOrEmpty(_formatCode)) FormatCode = FormatCode;
                SetSourceElement();
            }
        }

        private string GetNumLit()
        {
            var v = "";
            foreach (XmlNode node in _sourceElement.ChildNodes)
            {
                if(node.LocalName=="pt")
                {
                    v += node.FirstChild.InnerText + ",";
                }
            }
            if(v.Length>0)
            {
                v = "{" + v.Substring(0, v.Length - 1) + "}";
            }
            return v;
        }

        private void SetSourceElement()
        {
            var node = GetNode(_path);
            if(node!=null && node.HasChildNodes)
            {
                _sourceElement = (XmlElement)node.FirstChild;
            }
        }

        private void CreateNumLit(string value)
        {
            var nums = value.Substring(1, value.Length - 2).Split(',');
            if(nums.Length>0)
            {
                SetXmlNodeString($"{_path}/c:numLit/c:ptCount/@val", nums.Length.ToString(CultureInfo.InvariantCulture));
                var litNode = (XmlElement)GetNode($"{_path}/c:numLit");
                var idx = 0;
                foreach (var num in nums)
                {
                    var child = CreateLit(num.Trim(), idx++);
                    litNode.AppendChild(child);
                }
            }
        }

        private XmlElement CreateLit(string num, int idx)
        {
            XmlElement ptNode = TopNode.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
            ptNode.SetAttribute("idx",idx.ToString(CultureInfo.InvariantCulture));
            var vNode = TopNode.OwnerDocument.CreateElement("c", "v", ExcelPackage.schemaChart);
            vNode.InnerText = num;
            ptNode.AppendChild(vNode);
            return ptNode;
        }

        string _formatCode = "";
        /// <summary>
        /// The format code for the numeric source
        /// </summary>
        public string FormatCode
        {
            get
            {
                return _formatCode;
            }
            set
            {
                if (_sourceElement != null)
                {
                    switch (_sourceElement.LocalName)
                    {
                        case "numLit":
                            SetXmlNodeString(_path + "/c:numLit/c:formatCode", value);
                            break;
                        case "numRef":
                            SetXmlNodeString(_path + "/c:numRef/c:numCache/c:formatCode", value);
                            break;
                    }
                }
                _formatCode=value;
            }
        }
    }
}

