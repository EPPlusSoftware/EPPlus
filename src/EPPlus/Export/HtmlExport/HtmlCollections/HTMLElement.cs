using System.Collections.Generic;
using System.Xml.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Export.HtmlExport.HtmlCollections
{
    internal class HTMLElement
    {
        internal readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        internal List<HTMLElement> _childElements = new List<HTMLElement>();

        internal string ElementName { get; set; }

        internal string Content { get; set; }

        internal HTMLElement(string elementName)
        {
            ElementName = elementName;
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
