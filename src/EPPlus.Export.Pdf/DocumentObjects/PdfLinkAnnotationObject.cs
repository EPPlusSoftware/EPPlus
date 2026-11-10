using EPPlus.Export.Pdf.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        internal override string RenderDictionary()
        {
            return Body();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, Body());
        }
    }
}
