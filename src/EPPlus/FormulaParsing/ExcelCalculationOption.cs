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
            PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel;
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
                        PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel;
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
        /// Expressions in the formula calculation will be cached, to be reduced. 
        /// This increases speed, if having multiple formulas using the same expressions. 
        /// Caching increases memory consumption on calculate.
        /// </summary>
        public bool CacheExpressions { get; set; } = true;
        /// <summary>
        /// In some functions EPPlus will round double values to 15 significant figures before the value is handled. This is an option for Excel compatibility.
        /// </summary>
        public PrecisionAndRoundingStrategy PrecisionAndRoundingStrategy { get; set; }
        /// <summary>
        /// If true, EPPlus will calculate the cells in order calculating any dependent cells. Default.
        /// If false, EPPlus will calculate the cells without calculating dependent cells.
        /// </summary>
        public bool FollowDependencyChain
        {
            get;
            set;
        } = true;
        /// <summary>
        /// If true, EPPlus will download the images in the IMAGE function even if they exists in the package. The same URL will only be downloaded once.
        /// If false(default), EPPlus will only download images that doesn't exist in the package.
        /// </summary>
        public bool AlwaysRefreshImageFunction
        {
            get; 
            set; 
        } = false;

        /// <summary>
        /// Enables Unicode-aware string operations, ensuring correct handling of surrogate pairs for comparisons, substrings, and sorting within the library.
        /// </summary>
        public bool EnableUnicodeAwareStringOperations
        {
            get; set;
        } = false;
    }
}
