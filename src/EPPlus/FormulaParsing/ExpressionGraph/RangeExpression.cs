using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class RangeExpression : GroupExpression
    {
        public RangeExpression(ParsingContext ctx) : base(false,ctx)
        {

        }
        
        public override CompileResult Compile()
        {
            if (_result == null)
            {
                for (int i = 0; i < Children.Count - 1; i++)
                {
                    if (Children[i].Operator == null) return CompileResult.Empty;
                    _result = Children[i].Operator.Apply(_result ?? Children[i].Compile(), Children[i + 1].Compile(), Context);
                }
            }
            return _result;
        }
        public bool NeedsCalculation 
        { 
            get
            {
                foreach (var child in Children)
                {
                    if (child.ExpressionType == ExpressionType.Function)
                    {
                        return true;
                    }
                    if (child.ExpressionType == ExpressionType.NameValue)
                    {
                        var name = (NamedValueExpression)child;
                    }
                }
                return false;
            }
        }
        internal override ExpressionType ExpressionType => ExpressionType.RangeAddress;
    }
}
