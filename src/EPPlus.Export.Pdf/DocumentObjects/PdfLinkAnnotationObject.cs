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
using EPPlus.Export.Pdf.Helpers;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfLinkAnnotationObject : PdfObject
    {
        internal double LowerLeftX, LowerLeftY, UpperRightX, UpperRightY;
        internal string Uri;
        internal string ToolTip;
        internal bool IsGoTo;
        internal int DestPageObj;
        internal double DestX, DestY;

        public PdfLinkAnnotationObject(int objectNumber, double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY, string uri, string toolTip, int version = 0)
            : base(objectNumber, version)
        {
            LowerLeftX = lowerLeftX;
            LowerLeftY = lowerLeftY;
            UpperRightX = upperRightX;
            UpperRightY = upperRightY;
            Uri = uri;
            ToolTip = toolTip;
        }

        public static PdfLinkAnnotationObject GoTo(int objectNumber, double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY, int destPageObj, double destX, double destY, string toolTip, int version = 0)
        {
            var a = new PdfLinkAnnotationObject(objectNumber, lowerLeftX, lowerLeftY, upperRightX, upperRightY, null, toolTip, version);
            a.IsGoTo = true; a.DestPageObj = destPageObj; a.DestX = destX; a.DestY = destY;
            return a;
        }

        private string Body()
        {
            string action = IsGoTo
                ? $"/A << /S /GoTo /D [ {DestPageObj.ToPdfStringF0()} 0 R /XYZ {DestX.ToPdfStringF4()} {DestY.ToPdfStringF4()} null ] >>"
                : $"/A << /S /URI /URI {Uri.ToPdfLiteral()} >>";

            string contents = string.IsNullOrEmpty(ToolTip) ? "" : $"\n   /Contents {ToolTip.ToPdfTextString()}";

            return $"<< /Type /Annot /Subtype /Link\n" +
                   $"   /Rect [ {LowerLeftX.ToPdfStringF4()} {LowerLeftY.ToPdfStringF4()} {UpperRightX.ToPdfStringF4()} {UpperRightY.ToPdfStringF4()} ]\n" +
                   $"   /Border [ 0 0 0 ]\n" +
                   $"   /F 4{contents}\n" +
                   $"   {action} >>";
        }
        protected override void WriteBody(PdfWriter w) => w.Text(Body());
    }
}
