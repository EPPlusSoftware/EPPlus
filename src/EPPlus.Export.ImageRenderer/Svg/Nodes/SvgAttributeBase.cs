using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Nodes
{
    internal class SvgAttributeBase : RenderAttribute
    {
        internal SvgAttributeBase(string name, string strValue = "") 
        {
           SetName(name);
            if (string.IsNullOrEmpty(strValue) == false)
            {
                SetValue(strValue);
            }
        }

        internal virtual void SetName(string newName)
        {
            Name = newName;
        }
        protected virtual void SetValue(string strValue)
        {
            Value = strValue;
        }
    }
}
