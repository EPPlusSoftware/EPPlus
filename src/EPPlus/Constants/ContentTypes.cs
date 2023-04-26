/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/14/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
namespace OfficeOpenXml.Constants
{
    internal class ContentTypes
    {
        //Workbook & Worksheet
        internal const string contentTypeWorkbookDefault = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        internal const string contentTypeWorkbookMacroEnabled = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        internal const string contentTypeSharedString = @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        internal const string contentTypeMetaData = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
        
        //Theme
        internal const string contentTypeThemeOverride = "application/vnd.openxmlformats-officedocument.themeOverride+xml";
        internal const string contentTypeTheme = @"application/vnd.openxmlformats-officedocument.theme+xml";
        //Chart
        internal const string contentTypeChartStyle = "application/vnd.ms-office.chartstyle+xml";
        internal const string contentTypeChartColorStyle = "application/vnd.ms-office.chartcolorstyle+xml";        

        //Pivottables
        internal const string contentTypePivotTable = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
        internal const string contentTypePivotCacheDefinition = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
        internal const string contentTypePivotCacheRecords = @"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

        //VBA
        internal const string contentTypeVBA = @"application/vnd.ms-office.vbaProject";
        internal const string contentTypeVBASignature = @"application/vnd.ms-office.vbaProjectSignature";
        internal const string contentTypeVBASignatureAgile = @"application/vnd.ms-office.vbaProjectSignatureAgile";
        internal const string contentTypeVBASignatureV3 = @"application/vnd.ms-office.vbaProjectSignatureV3";

        internal const string contentTypeExternalLink = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";

        //Drawing
        internal const string contentTypeVml = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        internal const string contentTypeControlProperties = "application/vnd.ms-excel.controlproperties+xml";
        internal const string contentTypeChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        internal const string contentTypeChartEx = "application/vnd.ms-office.chartex+xml";

        //Rich data
        internal const string contentTypeRichDataValue = "application/vnd.ms-excel.rdrichvalue+xml";
        internal const string contentTypeRichDataValueStructure = "application/vnd.ms-excel.rdrichvaluestructure+xml";
        internal const string contentTypeRichDataValueType = "application/vnd.ms-excel.rdrichvaluetypes+xml";
    }
}
