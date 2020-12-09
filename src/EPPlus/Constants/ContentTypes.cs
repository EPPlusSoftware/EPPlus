using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Constants
{
    internal class ContentTypes
    {
        internal const string contentTypeWorkbookDefault = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        internal const string contentTypeWorkbookMacroEnabled = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        internal const string contentTypeSharedString = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        internal const string contentTypeControlProperties = "application/vnd.ms-excel.controlproperties+xml";
        internal const string contentTypeChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        internal const string contentTypeChartEx = "application/vnd.ms-office.chartex+xml";
        internal const string contentTypeThemeOverride = "application/vnd.openxmlformats-officedocument.themeOverride+xml";

        internal const string contentTypeTheme = @"application/vnd.openxmlformats-officedocument.theme+xml";
        internal const string contentTypeChartStyle = "application/vnd.ms-office.chartstyle+xml";
        internal const string contentTypeChartColorStyle = "application/vnd.ms-office.chartcolorstyle+xml";

        //Pivottables
        internal const string contentTypePivotTable = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
        internal const string contentTypePivotCacheDefinition = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
        internal const string contentTypePivotCacheRecords = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

        //VBA
        internal const string contentTypeVBA = @"application/vnd.ms-office.vbaProject";
        internal const string contentTypeVBASignature = @"application/vnd.ms-office.vbaProjectSignature";
    }
}
