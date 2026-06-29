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
namespace EPPlus.Export.Pdf.DocumentObjects.Functions
{
    internal abstract class PdfFunction : PdfObject
    {
        //    Implemented    Function type
        //    [ ]            0 Sampled function
        //    [X]            2 Exponential interpolation function
        //    [X]            3 Stitching function
        //    [ ]            4 PostScript calculator function

        internal double[] Domain;
        internal double[] Range;

        public PdfFunction(int objectNumber, int version = 0) : base(objectNumber, version) { }
    }
}
