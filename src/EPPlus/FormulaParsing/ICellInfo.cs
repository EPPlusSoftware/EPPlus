/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           EPPlus 6
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Information and help methods about a cell
    /// </summary>
    public interface ICellInfo
    {
        /// <summary>
        /// Address
        /// </summary>
        string Address { get; }
        /// <summary>
        /// WorksheetName
        /// </summary>
        string WorksheetName { get; }
        /// <summary>
        /// Row
        /// </summary>
        int Row { get; }
        /// <summary>
        /// Column
        /// </summary>
        int Column { get; }
        /// <summary>
        /// Id
        /// </summary>
        ulong Id { get; }
        /// <summary>
        /// Formula
        /// </summary>
        string Formula { get; }
        /// <summary>
        /// Value
        /// </summary>
        object Value { get; }
        /// <summary>
        /// Value double
        /// </summary>
        double ValueDouble { get; }
        /// <summary>
        /// Value double logical
        /// </summary>
        double ValueDoubleLogical { get; }
        /// <summary>
        /// Is hidden row
        /// </summary>
        bool IsHiddenRow { get; }
        /// <summary>
        /// Is excel error
        /// </summary>
        bool IsExcelError { get; }
        /// <summary>
        /// Tokens
        /// </summary>
        IList<Token> Tokens { get; }
    }
}
