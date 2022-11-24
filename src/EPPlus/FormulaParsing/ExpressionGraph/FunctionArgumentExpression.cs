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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class FunctionArgumentExpression : GroupExpression
    {
        private readonly Expression _function;

        public FunctionArgumentExpression(Expression function, ParsingContext ctx)
            : base(false, ctx)
        {
            _function = function;
            _parent = (ExpressionWithParent)function;
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override bool IgnoreCircularReference 
        { 
            get => base.IgnoreCircularReference; 
            set
            {
                base.IgnoreCircularReference = value;
                foreach(var childExpression in Children)
                {
                    childExpression.IgnoreCircularReference = value;
                }
            }
        }

        public override Expression PrepareForNextChild()
        {
            return _function.PrepareForNextChild();
        }
        internal override ExpressionType ExpressionType => ExpressionType.FunctionArgument;
        internal ExcelFunction Function
        {
            get
            {
                return Context.Configuration.FunctionRepository.GetFunction(_function.ExpressionString);
            }
        }
    }
}
