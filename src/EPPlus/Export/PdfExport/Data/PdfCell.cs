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
using EPPlus.Export.Pdf.Layout;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfCell : PdfCellBase
    {
        public PdfCellStyle CellStyle;

        public bool Merged;
        public PdfCell Main;
        public ExcelAddressBase MergedAddress;
    }
}
