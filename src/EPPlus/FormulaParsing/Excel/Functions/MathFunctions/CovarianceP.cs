/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/12/2020         EPPlus Software AB       Version 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.5",
        Description = "Returns covariance, the average of the products of deviations for each data point pair in two data sets.")]
    internal class CovarianceP : Covar
    {
        public override string NamespacePrefix => "_xlfn.";
    }
}
