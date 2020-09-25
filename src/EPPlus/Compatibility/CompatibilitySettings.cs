/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
 using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace OfficeOpenXml.Compatibility
{
    /// <summary>
    /// Settings to stay compatible with older versions of EPPlus
    /// </summary>
    public class CompatibilitySettings
    {
        private ExcelPackage excelPackage;


        internal CompatibilitySettings(ExcelPackage excelPackage)
        {
            this.excelPackage = excelPackage;
        }
#if Core
        /// <summary>
        /// If the worksheets collection of the ExcelWorkbook class is 1 based.
        /// This property can be set from appsettings.json file.
        /// <code>
        ///     {
        ///       "EPPlus": {
        ///         "ExcelPackage": {
        ///           "Compatibility": {
        ///             "IsWorksheets1Based": true //Default and recommended value is false
        ///           }
        ///         }
        ///       }
        ///     }
        /// </code>
        /// </summary>
#else
        /// <summary>
        /// If the worksheets collection of the ExcelWorkbook class is 1 based.
        /// This property can be set from app.config file.
        /// <code>
        ///   <appSettings>
        ///    <!--Set worksheets collection to start from one.Default is 0. Set to true for backward compatibility reasons only!-->  
        ///    <add key = "EPPlus:ExcelPackage.Compatibility.IsWorksheets1Based" value="true" />
        ///   </appSettings>
        /// </code>
        /// </summary>
#endif

        public bool IsWorksheets1Based
        {
            get
            {
                return excelPackage._worksheetAdd==1;
            }
            set
            {
                excelPackage._worksheetAdd = value ? 1 : 0;
                if(excelPackage._workbook!=null && excelPackage._workbook._worksheets!=null)
                {
                    excelPackage.Workbook.Worksheets.ReindexWorksheetDictionary();

                }
            }
        }
   }
}
