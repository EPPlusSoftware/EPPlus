/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/31/2022         EPPlus Software AB           EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// EPPlus implementation of the <see cref="INameInfo"/> interface
    /// </summary>
    public class NameInfo : INameInfo
    {
        /// <summary>
        /// Id
        /// </summary>
        public ulong Id { get; set; }
        /// <summary>
        /// Worksheet name
        /// </summary>
        public int wsIx { get; set; }
        /// <summary>
        /// The name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Formula of the name
        /// </summary>
        public string Formula { get; set; }
        /// <summary>
        /// Tokens
        /// </summary>
        public IList<Token> Tokens { get; internal set; }
        /// <summary>
        /// Value
        /// </summary>
        public object Value { get; set; }
    }
}
