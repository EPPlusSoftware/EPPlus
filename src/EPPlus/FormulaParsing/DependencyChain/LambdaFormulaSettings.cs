/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/03/2025         EPPlus Software AB       EPPlus 8.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal class LambdaFormulaSettings
    {
        public HashSet<int> LambdaTokens { get; private set; }

        internal Stack<int> LambdaStackNumbers { get; private set; } = new Stack<int>();

        internal Stack<short> NumberOfLambdaVariables { get; set; } = new Stack<short>();

        internal Stack<LambdaExpressionStackPosition> CurrentLambdaExpressions { get; set; } = new Stack<LambdaExpressionStackPosition>();

        internal Stack<short> LambdaArgsAdded { get; set; } = new Stack<short>();

        internal Stack<short> InvokeLambdaAt { get; set; } = new Stack<short>();

        internal short CurrentLambdaArgsAdded
        {
            get
            {
                if (LambdaArgsAdded.Count == 0) return 0;
                return LambdaArgsAdded.Peek();
            }
        }

        public void AddLambdaToken(int tokenIx)
        {
            LambdaTokens ??= [];
            if (LambdaTokens.Contains(tokenIx)) return;
            LambdaTokens.Add(tokenIx);
        }

        internal void Reset()
        {
            LambdaTokens?.Clear();
            LambdaStackNumbers.Clear();
            NumberOfLambdaVariables.Clear();
            LambdaArgsAdded.Clear();
            InvokeLambdaAt.Clear();
            CurrentLambdaExpressions.Clear();
        }
    }
}
