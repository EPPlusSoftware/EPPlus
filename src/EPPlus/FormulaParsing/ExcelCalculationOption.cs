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
using System.IO;
#else
using System.Configuration;
#endif


namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Options used by the formula parser
    /// </summary>
    public class ExcelCalculationOption
    {
        public ExcelCalculationOption()
        {
            AllowCircularReferences = false;
#if (Core)
            var build = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", true, false);
            var c = build.Build();

            var configValue = c["EPPlus:ExcelPackage:AllowCircularReferences"];
            if(bool.TryParse(configValue, out bool allow))
            {
                AllowCircularReferences = allow;
            }

#else
            if(bool.TryParse(ConfigurationManager.AppSettings["EPPlus:ExcelPackage.AllowCircularReferences"], out bool allow))
            {
                AllowCircularReferences = allow;
            }
#endif
        }
        /// <summary>
        /// Do not throw an exception if the formula parser encounters a circular reference
        /// </summary>
        public bool AllowCircularReferences { get; set; }
    }
}
