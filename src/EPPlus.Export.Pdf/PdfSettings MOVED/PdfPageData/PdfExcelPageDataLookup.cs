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
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfSettings.PdfPageData
{
    internal class PdfExcelPageDataLookup
    {
        //Not to be used but kept for reference, This table has data on how many rows and columns each font/theme gives in an A4 page.
        internal static Dictionary<string, int[]> PdfExcelA4PageData = new Dictionary<string, int[]>()
        {
            { "Aptos",                  [48,    9] },
            { "Aptos Display",          [48,    9] },
            { "Aptos Narrow",           [48,    9] },
            { "Arial",                  [53,    8] },
            { "Arial Black",            [42,    7] },
            { "Arial Narrow",           [53,   10] },
            { "Bookman Old Style",      [53,    8] },
            { "Calibri",                [50,    9] },
            { "Calibri Light",          [50,    9] },
            { "Calisto MT",             [53,    9] },
            { "Cambria",                [53,    8] },
            { "Candara",                [50,    8] },
            { "Century Gothic",         [51,    8] },
            { "Century Schoolbook",     [51,    8] },
            { "Consolas",               [52,    8] },
            { "Constantia",             [50,    9] },
            { "Corbel",                 [50,    9] },
            { "Courier New",            [57,    8] },
            { "Franklin Gothic Book",   [49,    8] },
            { "Franklin Gothic Medium", [49,    8] },
            { "Garamond",               [55,   10] },
            { "Georgia",                [54,    8] },
            { "Gill Sans MT",           [44,    9] },
            { "Impact",                 [51,    9] },
            { "Liberation Serif",       [53,    9] },
            { "MS Gothic",              [59,    9] },
            { "Palatino Linotype",      [44,    9] },
            { "Rockwell",               [54,    9] },
            { "Rockwell Condensed",     [52,   12] },
            { "SegoeUI",                [44,    9] },
            { "Tahoma",                 [51,    9] },
            { "Times New Roman",        [52,    9] },
            { "Trebuchet MS",           [49,    9] },
            { "Tw Cen MT",              [54,    8] },
            { "Verdana",                [50,    8] },
        };
    }
}