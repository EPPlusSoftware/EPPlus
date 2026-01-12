using EPPlus.Export.ImageRenderer.Svg.Nodes;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.NodeAttributes
{
    internal class SvgElement
    {
        internal readonly List<SvgAttributeBase> _attributes = new List<SvgAttributeBase>();

        internal List<SvgElement> _childElements = new List<SvgElement>();

        internal string ElementName { get; set; }

        internal string Content { get; set; }

        /// <summary>
        /// If true element cannot ever have content
        /// </summary>
        internal bool IsVoidElement { get; private set; }

        internal SvgElement(string elementName)
        {
            ElementName = elementName;
        }

        /// <summary>
        /// Attempt to add attribute value of unknown type
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="attributeValue"></param>
        public void AddAttribute(string attributeName, object attributeValue)
        {
            if(attributeValue != null)
            {
                //var objStr = string.Format(CultureInfo.InvariantCulture, attributeValue.ToString());
                var objStr = ConvertUtil.GetValueForXml(attributeValue, true);
                AddAttribute(attributeName, objStr);
            }
        }

        /// <summary>
        /// Add attribute with a value
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="attributeValue"></param>
        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new SvgAttributeBase(attributeName, attributeValue));
        }

        /// <summary>
        /// Add attribute without it having a value
        /// </summary>
        /// <param name="attributeName"></param>
        public void AddAttributeValueLess(string attributeName)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            _attributes.Add(new SvgAttributeBase(attributeName));
        }

        /// <summary>
        /// Add child element
        /// </summary>
        /// <param name="element"></param>
        public void AddChildElement(SvgElement element)
        {
            _childElements.Add(element);
        }
    }
}
