using System;
using System.Collections.Generic;
namespace OfficeOpenXml.Constants
{
    internal class Schemas
    {
        //Main Schemas
        internal const string schemaMain = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal const string schemaMarkupCompatibility = @"http://schemas.openxmlformats.org/markup-compatibility/2006";

        //Dynamic arrays
        internal const string schemaDynamicArray = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

        // Richdata (used in worksheet.sortstate)
        internal const string schemaRichData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        internal const string schemaRichData2 = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
        internal const string schemaRichValueRel = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    }
}
