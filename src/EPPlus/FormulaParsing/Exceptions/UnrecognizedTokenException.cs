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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    /// <summary>
    /// Unrecognized token exception
    /// </summary>
    public class UnrecognizedTokenException : Exception
    {
        /// <summary>
        /// Constructor. Token exception
        /// </summary>
        /// <param name="token">Tje token that can not be recognized</param>
        public UnrecognizedTokenException(Token token)
            : base( "Unrecognized token: " + token.Value)
        {

        }
    }
}
