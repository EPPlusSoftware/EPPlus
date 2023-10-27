using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Determinator
{
    internal static class DxfStylChecker
    {
        internal static int GetIdIfShouldAdd(Dictionary<string, List<ExcelConditionalFormattingRule>> conditionalFormattings, StyleCache cache, string cellAddress)
        {
            if (cellAddress != null && conditionalFormattings.ContainsKey(cellAddress))
            {
                if (conditionalFormattings.ContainsKey(cellAddress))
                {

                }
            }
            return -1;
        }

        //internal static int GetIdFromCache(ExcelDxfStyleConditionalFormatting dxfs, StyleCache cache)
        //{
        //    if (dxfs != null)
        //    {
        //        if (!IsAddedToCache(dxfs, cache, out int id))
        //        {
        //            return id;
        //        }
        //    }

        //    return -1;
        //}
    }
}
