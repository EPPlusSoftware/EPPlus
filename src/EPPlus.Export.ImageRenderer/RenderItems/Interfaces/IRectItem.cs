using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    internal interface IRectItem
    {
        double Top { get; set; }
        double Left { get; set; }
        double Height { get; set; }
        double Width { get; set; }
    }
}
