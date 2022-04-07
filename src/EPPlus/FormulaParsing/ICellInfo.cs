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
        string Address { get; }

        string WorksheetName { get; }
        int Row { get; }
        int Column { get; }

        ulong Id { get; }
        string Formula { get; }
        object Value { get; }
        double ValueDouble { get; }
        double ValueDoubleLogical { get; }
        bool IsHiddenRow { get; }
        bool IsExcelError { get; }
        IList<Token> Tokens { get; }
    }
}
