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
using System.Collections.Generic;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
    internal class FormulaCell
    {
        internal int Index { get; set; }
        /// <summary>
        /// NOTE: This is the position in the ExcelWorksheets._worksheets collection. Cannot be used direcly with Worksheets[] indexer.
        /// </summary>
        internal int wsIndex { get; set; }
        internal int Row { get; set; }
        internal int Column { get; set; }
        internal string Formula { get; set; }

        internal string CircularRefAddress { get; set; }
        internal List<Token> Tokens { get; set; }
        internal int tokenIx = 0;
        internal int addressIx = 0;
        internal CellStoreEnumerator<object> iterator;
        internal ExcelWorksheet iteratorWs;
        internal ExcelWorksheet ws;
    }
}