using OfficeOpenXml.PDF.PdfSettings;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects.PdfShadings
{
    public enum DeviceColorSpace
    {
        DeviceGray,
        DeviceRGB,
        DeviceCMYK,
    }

    //    Implemented    Function type
    //    [ ]            1 Function-based shading
    //    [X]            2 Axial shading
    //    [ ]            3 Radial shading
    //    [ ]            4 Free-form Gouraud-shaded triangle mesh
    //    [ ]            5 Lattice-form Gouraud-shaded triangle mesh
    //    [ ]            6 Coons patch mesh
    //    [ ]            7 Tensor-product patch mesh

    internal class PdfShading : PdfObject
    {
        internal DeviceColorSpace ColorSpace;
        internal double[] Background = null;
        internal PdfRect BBox = null;
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
    }
}
