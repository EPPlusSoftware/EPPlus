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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// NameInfo
    /// </summary>
    public interface INameInfo
    {
        /// <summary>
        /// Id
        /// </summary>
        ulong Id { get; }
        /// <summary>
        /// wsIx
        /// </summary>
        int wsIx { get; }
        /// <summary>
        /// Name
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Formula
        /// </summary>
        string Formula { get; }
        /// <summary>
        /// Value
        /// </summary>
        object Value { get; }
        /// <summary>
        /// IsRelative
        /// </summary>
        bool IsRelative { get; }
        /// <summary>
        /// Get relative formula
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        string GetRelativeFormula(int row, int col);
        /// <summary>
        /// Get relative range
        /// </summary>
        /// <param name="ri"></param>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        IRangeInfo GetRelativeRange(IRangeInfo ri, FormulaCellAddress currentCell);
        /// <summary>
        /// Get value
        /// </summary>
        /// <param name="currentCell"></param>
        /// <returns></returns>
        object GetValue(FormulaCellAddress currentCell);
    }
}
