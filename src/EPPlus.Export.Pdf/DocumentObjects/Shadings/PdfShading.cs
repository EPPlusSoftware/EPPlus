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
using EPPlus.Graphics;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Shadings
{
    internal enum DeviceColorSpace
    {
        DeviceGray,
        DeviceRGB,
        DeviceCMYK,
    }

    internal class PdfShading : PdfObject
    {
        //    Implemented    Function type
        //    [ ]            1 Function-based shading
        //    [X]            2 Axial shading
        //    [ ]            3 Radial shading
        //    [ ]            4 Free-form Gouraud-shaded triangle mesh
        //    [ ]            5 Lattice-form Gouraud-shaded triangle mesh
        //    [ ]            6 Coons patch mesh
        //    [ ]            7 Tensor-product patch mesh

        internal DeviceColorSpace ColorSpace;
        internal double[] Background = null;
        internal Rect BBox = null;
        internal bool? AntiAlias = null;

        public PdfShading(int objectNumber, int version = 0) : base(objectNumber, version) { }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Shading\n" +
                            $"   /ShadingType 0\n" +
                            $"   /ColorSpace {ColorSpace.ToString()}");
            if (Background != null)
            {
                //add background
            }
            if (BBox != null)
            {
                //add bbox
            }
            if (AntiAlias != null)
            {
                sb.AppendFormat($"\n   /AntiAlias {AntiAlias.ToString()}");
            }
            return sb.ToString();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Shading\n" +
                            $"   /ShadingType 0\n" +
                            $"   /ColorSpace {ColorSpace.ToString()}");
            if (Background != null)
            {
                //add background
            }
            if (BBox != null)
            {
                //add bbox
            }
            if (AntiAlias != null)
            {
                sb.AppendFormat($"\n   /AntiAlias {AntiAlias.ToString()}");
            }
            WriteAscii(bw, sb.ToString());
        }
    }
}
