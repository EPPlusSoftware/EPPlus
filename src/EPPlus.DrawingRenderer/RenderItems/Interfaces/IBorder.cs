using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Interfaces
{
    public interface IBorder : IFill
    {
        double? BorderWidth { get; set; }
        double[] BorderDashArray { get; set; }
        double? BorderDashOffset { get; set; }
    }
}
