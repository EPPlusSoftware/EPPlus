/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using System.Collections.Generic;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Export.HtmlExport.HtmlCollections
{
    internal class HTMLElement
    {
        internal readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        internal List<HTMLElement> _childElements = new List<HTMLElement>();

        internal string ElementName { get; set; }

        internal string Content { get; set; }

        internal bool IsVoidElement { get; private set; }

        internal HTMLElement(string elementName)
        {
            ElementName = elementName;
            if(HtmlElements.VoidElements.Contains(elementName)) 
            {
                IsVoidElement = true;
            }
        }

        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }

        public void AddChildElement(HTMLElement element)
        {
            _childElements.Add(element);
        }
    }
}
