/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  19/3/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions
{
    internal class GroupByBaseArgs
    {
        public IRangeInfo RowFields { get; set; }
        public IRangeInfo Values { get; set; }
        public LambdaCalculator Function { get; set; }
        public List<LambdaCalculator> Functions { get; set; } = new List<LambdaCalculator>();
        public FunctionLayout FunctionLayout { get; set; } = FunctionLayout.Single;
        public FieldHeaders Headers { get; set; } = FieldHeaders.Missing;
        public int TotalDepth { get; set; } = 1;
        public int[] SortOrders { get; set; } = new[] { 1 };
        public IRangeInfo FilterArray { get; set; } = null;
        public FieldRelationship FieldRelationship { get; set; } = FieldRelationship.Hierarchy;
        public List<object[]> AllValuesInOrder { get; set; } = new List<object[]>();
    }
}
