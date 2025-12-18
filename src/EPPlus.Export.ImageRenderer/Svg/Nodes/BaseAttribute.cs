using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.NodeAttributes
{
    internal abstract class BaseAttribute
    {
       public string Name { get; internal protected set; }

       internal protected string Value { get; protected set; }
    }
}
