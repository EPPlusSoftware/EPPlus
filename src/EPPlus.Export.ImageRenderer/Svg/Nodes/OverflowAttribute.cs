using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Nodes
{
    enum eOverFlowValues
    {
        Visible,
        Hidden,
        Scroll,
        Auto
    }

    internal class OverflowAttribute : SvgAttributeBase
    {
        //Default for attribute is visible: https://developer.mozilla.org/en-US/docs/Web/SVG/Reference/Attribute/overflow
        private eOverFlowValues _overflowVal;

        internal eOverFlowValues OverFlowValue
        {
            get { return _overflowVal; }
            set 
            {
                switch (value)
                {
                    case eOverFlowValues.Visible:
                        Value = "Visible";
                        break;
                    case eOverFlowValues.Hidden:
                        Value = "Hidden";
                        break;
                    case eOverFlowValues.Scroll:
                        Value = "Scroll";
                        break;
                    case eOverFlowValues.Auto:
                        Value = "Auto";
                        break;
                }
                _overflowVal = value;
            }
        }

        internal OverflowAttribute(eOverFlowValues overflowValue = eOverFlowValues.Visible) : base("Overflow")
        {
            OverFlowValue = overflowValue;
        }

        internal override void SetName(string newName)
        {
            throw new System.Exception("A Type Specific Attribute cannot be set to a different name!!");
        }
    }
}
