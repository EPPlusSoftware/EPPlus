using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    public interface IFill
    {
        string Color { get; set; }
        double? Opacity { get; set; }
        PathFillMode FillMode { get; set; }
    }
}
