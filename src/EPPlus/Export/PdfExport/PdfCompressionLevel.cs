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
namespace OfficeOpenXml.Export.PdfExport
{
    /// <summary>
    /// Controls the compression applied to PDF content streams and images.
    /// More compression means a smaller file but a slower export.
    /// </summary>
    public enum PdfCompressionLevel
    {
        /// <summary>Fastest export, largest file.</summary>
        Fastest,
        /// <summary>A balance between export speed and file size.</summary>
        Balanced,
        /// <summary>Smallest file, slowest export. This is the default.</summary>
        Smallest
    }
}
