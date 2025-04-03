using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{
    internal class LambdaFormulaSettings
    {
        public HashSet<int> LambdaTokens { get; private set; }

        internal Stack<int> LambdaStackNumbers { get; private set; } = new Stack<int>();

        internal Stack<short> NumberOfLambdaVariables { get; set; } = new Stack<short>();

        internal Stack<LambdaExpressionStackPosition> CurrentLambdaExpressions { get; set; } = new Stack<LambdaExpressionStackPosition>();

        internal Stack<short> LambdaArgsAdded { get; set; } = new Stack<short>();

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
    }
}
