using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Nodes
{
    internal class BaseElement
    {
        internal readonly List<BaseAttribute> _attributes = new List<BaseAttribute>();

        internal string ElementName { get; set; }

        internal string Content { get; set; }

        //internal B
    }
}
