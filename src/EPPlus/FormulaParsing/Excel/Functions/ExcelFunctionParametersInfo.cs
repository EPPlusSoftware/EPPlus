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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Function parameters info
    /// </summary>
    public class ExcelFunctionParametersInfo
    {
        /// <summary>
        /// Default
        /// </summary>
        public static ExcelFunctionParametersInfo Default 
        { 
            get
            {
                return new ExcelFunctionParametersInfo();
            }
        }
        Func<int, FunctionParameterInformation> _getParameter=null;
        private ExcelFunctionParametersInfo()
        {

        }
        /// <summary>
        /// Constructor getParameter
        /// </summary>
        /// <param name="getParameter"></param>
        public ExcelFunctionParametersInfo(Func<int, FunctionParameterInformation> getParameter)
        {
            _getParameter = getParameter;
        }
        /// <summary>
        /// Has normal arguments
        /// </summary>
        public bool HasNormalArguments
        {
            get 
            { 
                return _getParameter == null;
            }
        }
        /// <summary>
        /// Get information about the parameter at the position at <paramref name="argumentIndex"/>
        /// </summary>
        /// <param name="argumentIndex">The position of the parameter</param>
        /// <returns>The parameter informations</returns>
        public virtual FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if(_getParameter== null)
            {
                return FunctionParameterInformation.Normal;
            }
            return _getParameter(argumentIndex);
        }
    }
}
