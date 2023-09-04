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
    public class ExcelFunctionParametersInfo
    {
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
        public ExcelFunctionParametersInfo(Func<int, FunctionParameterInformation> getParameter)
        {
            _getParameter = getParameter;
        }
        public bool HasNormalArguments
        {
            get 
            { 
                return _getParameter == null;
            }
        }
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
