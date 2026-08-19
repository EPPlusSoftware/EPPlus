using EPPlus.Export.Pdf.DocumentObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfImageResource : PdfResource
    {
        internal int objectNumber;
        internal readonly byte[] ImageBytes;

        public PdfImageResource(int labelNumber, byte[] imageBytes)
            : base("Im", labelNumber)
        {
            ImageBytes = imageBytes;
        }

        public PdfImageXObject GetImageObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
            return new PdfImageXObject(objectNumber, ImageBytes, version);
        }
    }
}
