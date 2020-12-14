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
    internal class ExtLstUris
    {
        //Legacy Object Wrapper
        internal const string LegacyObjectWrapperUri = "{63B3BB69-23CF-44E3-9099-C40C66FF867C}";

        //Pivot Table
        internal const string PivotTableDefinitionUri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}";
        internal const string PivotTableDataUri = "{44433962-1CF7-4059-B4EE-95C3D5FFCF73}";
        internal const string PivotTableServerFormatsUri = "{C510F80B-63DE-4267-81D5-13C33094786E}";
        internal const string PivotTableUISettingsUri = "{E67621CE-5B39-4880-91FE-76760E9C1902}";
        internal const string PivotTableDefinition16Uri = "{747A6164-185A-40DC-8AA5-F01512510D54}";

        //Pivot Table Cache Defintion
        internal const string PivotCacheDefinitionUri = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}";           //Excel require lower case on 48be???
        internal const string TimelinePivotCacheDefinitionUri = "{5DA0FC9A-693D-419c-AD59-312A39285967}";
        internal const string PivotCacheIdVersionUri = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";

        //Slicer
        internal const string SlicerCachePivotTablesUri = "{03082B11-2C62-411c-B77F-237D8FCFBE4C}";
        internal const string TableSlicerCacheUri = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";
        internal const string SlicerCacheHideItemsWithNoDataUri = "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}";

        //Slicer in worksheet
        internal const string WorkbookSlicerPivotTableUri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";
        internal const string WorkbookSlicerTableUri = "{46BE6895-7355-4a93-B00E-2C351335B9C9}";
        
        //Slicer in worksheet
        internal const string WorksheetSlicerPivotTableUri = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";        
        internal const string WorksheetSlicerTableUri = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";
    }
}

