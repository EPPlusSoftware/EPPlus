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
#if (Core)
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
#else
using System.Configuration;
using System.Collections.Generic;
#endif


namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Options used by the formula parser
    /// </summary>
    public class ExcelCalculationOption
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelCalculationOption()
        {
            AllowCircularReferences = false;
            PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.DotNet;
            var initErrors = new List<ExcelInitializationError>();

#if (Core)
            var configValue = ExcelConfigurationReader.GetJsonConfigValue("EPPlus:ExcelPackage:AllowCircularReferences", ExcelPackage.GlobalConfiguration, initErrors);
            if(bool.TryParse(configValue, out bool allow))
            {
                AllowCircularReferences = allow;
            }
            //var roundingStrategy = c["EPPlus:ExcelPackage:PrecisionAndRoundingStrategy"];
            var roundingStrategy = ExcelConfigurationReader.GetJsonConfigValue("EPPlus:ExcelPackage:PrecisionAndRoundingStrategy", ExcelPackage.GlobalConfiguration, initErrors);
            if (Enum.TryParse(roundingStrategy, out PrecisionAndRoundingStrategy precisionAndRoundingStrategy))
            {
                PrecisionAndRoundingStrategy = precisionAndRoundingStrategy;
            }

#else
            var acr = ExcelConfigurationReader.GetValueFromAppSettings("EPPlus:ExcelPackage.AllowCircularReferences", ExcelPackage.GlobalConfiguration, initErrors);
            if(bool.TryParse(acr, out bool allow))
            {
                AllowCircularReferences = allow;
            }
            // no Enum.TryParse in .NET 35...
            var roundingStrategy = ExcelConfigurationReader.GetValueFromAppSettings("EPPlus:ExcelPackage.PrecisionAndRoundingStrategy", ExcelPackage.GlobalConfiguration, initErrors);
            if(!string.IsNullOrEmpty(roundingStrategy))
            {
                switch(roundingStrategy.ToLower())
                {
                    case "dotnet":
                        PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.DotNet;
                        break;
                    case "excel":
                        PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel;
                        break;
                    default:
                        PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.DotNet;
                        break;
                }
            }
#endif
        }
        /// <summary>
        /// Do not throw an exception if the formula parser encounters a circular reference
        /// </summary>
        public bool AllowCircularReferences { get; set; }

        /// <summary>
        /// In some functions EPPlus will round double values to 15 significant figures before the value is handled. This is an option for Excel compatibility.
        /// </summary>
        public PrecisionAndRoundingStrategy PrecisionAndRoundingStrategy { get; set; }
    }
}
