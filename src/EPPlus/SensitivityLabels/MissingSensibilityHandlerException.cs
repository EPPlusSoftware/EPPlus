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
#if(!NET35)
#endif
namespace OfficeOpenXml.SensitivityLabels
{
    /// <summary>
    /// Thown if no sensibility handler is set, and a package is encrypted with a protected sensibility label.
    /// </summary>
    [Serializable]    
    public class MissingSensibilityHandlerException : Exception
    {
        internal MissingSensibilityHandlerException()
        {
        }

        internal MissingSensibilityHandlerException(string message) : base(message)
        {
        }

        internal MissingSensibilityHandlerException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
