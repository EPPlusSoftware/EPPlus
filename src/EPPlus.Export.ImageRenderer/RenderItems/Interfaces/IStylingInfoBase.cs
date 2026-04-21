using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    internal interface IStylingInfoBase
    {
        string FillColor { get; set; }
        string BorderColor { get; set; }
    }
}
