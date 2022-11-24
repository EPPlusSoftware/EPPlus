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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public abstract class Expression
    {
        internal string ExpressionString { get; private set; }
        private readonly List<Expression> _children = new List<Expression>();
        protected ParsingContext Context { get; private set; }
        public IList<Expression> Children { get { return _children; } }
        public IOperator Operator { get; set; }
        internal abstract ExpressionType ExpressionType { get; }
        public abstract bool IsGroupedExpression { get; }
        /// <summary>
        /// If set to true, <see cref="ExcelAddressExpression"></see>s that has a circular reference to their cell will be ignored when compiled
        /// </summary>
        public virtual bool IgnoreCircularReference
        {
            get; set;
        }

        public Expression(ParsingContext ctx)
        {
            Context = ctx;
        }

        public Expression(string expression, ParsingContext ctx)
        {
            ExpressionString = expression;
            Operator = null;
            Context = ctx;
        }

        public virtual bool HasChildren
        {
            get { return _children.Count>0; }
        }

        /// <summary>
        /// Prepares the expression for next child expression
        /// </summary>
        /// <returns></returns>
        public virtual Expression  PrepareForNextChild()
        {
            return this;
        }

        /// <summary>
        /// Prepares the expression for next child expression.
        /// </summary>
        /// <param name="token"><see cref="Token"/> that is relevant in the context.</param>
        /// <returns></returns>
        public virtual Expression PrepareForNextChild(Token token)
        {
            return this;
        }

        /// <summary>
        /// Adds a child expression.
        /// </summary>
        /// <param name="child">The child expression to add</param>
        /// <returns></returns>
        public virtual Expression AddChild(Expression child)
        {
            if (_children.Any())
            {
                var last = _children.Last();
                //child.Prev = last;
                //last.Next = child;
            }
            _children.Add(child);
            return child;
        }

        public virtual Expression MergeWithNext(IList<Expression> expressions, int index)
        {
            var expression = this;
            var Next = GetItem(expressions, index + 1);
            var Prev = GetItem(expressions, index - 1);
            if (Next != null && Operator != null)
            {
                var left = Compile();
                var right = Next.Compile();
                var result = Operator.Apply(left, right, Context);
                expression = ExpressionConverter.GetInstance(Context).FromCompileResult(result);
                if (expression is ExcelErrorExpression)
                {
                    //expression.Next = null;
                    //expression.Prev = null;
                    return expression;
                }
                if (Next != null)
                {
                    expression.Operator = Next.Operator;
                }
                else
                {
                    expression.Operator = null;
                }
                //expression.Next = Next.Next;
                //if (expression.Next != null) expression.Next.Prev = expression;
                //expression.Prev = Prev;
                expressions.Insert(index, expression);
                expressions.Remove(this);
                expressions.Remove(Next);
            }
            else
            {
                throw (new FormatException("Invalid formula syntax. Operator missing expression."));
            }
            //if (Prev != null)
            //{
            //    Prev.Next = expression;
            //}            
            return expression;
        }

        private Expression GetItem(IList<Expression> expressions, int index)
        {
            if(index < 0 || index >= expressions.Count)
            {
                return null;
            }
            return expressions[index];
        }

        internal abstract Expression Clone();
        internal virtual Expression Clone(int rowOffset, int colOffset)
        {
            return CloneExpressionWithOffset(this, rowOffset, colOffset);
        }
        protected Expression CloneExpressionWithOffset(Expression e, int rowOffset, int colOffset)
        {
            var clone = e.Clone();
            foreach(var c in e.Children)
            {
                clone.Children.Add(c.Clone(rowOffset, colOffset));
            }
            clone.Operator = Operator;
            clone._result = _result;
            return clone;
        }
        protected Expression CloneMe(Expression e)
        {
            //foreach (var c in Children)
            //{
            //    e.Children.Add(c.Clone());
            //}
            e.Operator = Operator;
            e._result = _result;
            return e;
        }
        protected CompileResult _result;
        public abstract CompileResult Compile();

    }
}
