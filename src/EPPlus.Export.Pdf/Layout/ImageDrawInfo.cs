using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Layout
{
    internal struct ImageDrawInfo
    {
        public double X;
        public double Y;
        public double Width;
        public double Height;
        public byte[] ImageBytes;
        public string Name;
    }
}
