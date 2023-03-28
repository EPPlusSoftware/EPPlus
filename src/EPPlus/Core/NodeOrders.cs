using System.Collections.Generic;
namespace OfficeOpenXml.Core
{
    internal static class NodeOrders
    {
        internal static readonly Dictionary<string, int> WorksheetTopElementOrder = new Dictionary<string, int>()
        {
            { "sheetPr",0 },
            { "dimension",1 },
            { "sheetViews",2 },
            { "sheetFormatPr",3 },
            { "cols",4 },
            { "sheetData",5 },
            { "sheetCalcPr",6 },
            { "sheetProtection",7 },
            { "protectedRanges",8 },
            { "scenarios",9 },
            { "autoFilter",10 },
            { "sortState",11 },
            { "dataConsolidate",12 },
            { "customSheetViews",13 },
            { "mergeCells",14 },
            { "phoneticPr",15 },
            { "conditionalFormatting",16 },
            { "dataValidations",17 },
            { "hyperlinks",18 },
            { "printOptions",19 },
            { "pageMargins",20 },
            { "pageSetup",21 },
            { "headerFooter",22 },
            { "rowBreaks",23 },
            { "colBreaks",24 },
            { "customProperties",25 },
            { "cellWatches",26 },
            { "ignoredErrors",27 },
            { "smartTags",28 },
            { "drawing",29 },
            { "drawingHF",30 },
            { "picture",31 },
            { "oleObjects",32 },
            { "controls",33 },
            { "webPublishItems",34 },
            { "tableParts",35 },
            { "extLst",36 },
        };
    }
}
