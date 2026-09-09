/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.DocumentObjects;

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
