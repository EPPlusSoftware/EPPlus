using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Style
{
    internal interface IStyleExportDrawing
    {
        string StyleKey { get; }

        bool HasStyle { get; }

        IFillBasic Fill { get; }
    }
}
