using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Nodes
{
    internal abstract class RenderAttribute : BaseAttribute
    {
        internal string Render()
        {
            //TODO; Make static const string to only replace minor values in?
            string renderedString = $" {Name}";
            if (string.IsNullOrEmpty(Value) == false)
            {
                renderedString += $"=\"{Value}\"";
            }
            return renderedString;
        }
    }
}
