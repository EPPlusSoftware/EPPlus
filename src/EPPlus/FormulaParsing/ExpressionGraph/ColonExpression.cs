using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ColonExpression : FunctionExpression
    {
        public ColonExpression(string expression, ParsingContext parsingContext) : base(expression, parsingContext, false)
        {
        }

        public override CompileResult Compile(bool treatEmptyAsZero = true)
        {
            var prevIsAddress = Prev.GetType() == typeof(ExcelAddressExpression);
            var prevIsOffset = Prev.GetType() == typeof(FunctionExpression) && ((FunctionExpression)Prev).ExpressionString.ToLower() == "offset";
            var nextIsAddress = Next.GetType() == typeof(ExcelAddressExpression);
            var nextIsOffset = Next.GetType() == typeof(FunctionExpression) && ((FunctionExpression)Next).ExpressionString.ToLower() == "offset";

            if (!prevIsAddress && !prevIsOffset) return new CompileResult(eErrorType.Value);
            if (!nextIsAddress && !nextIsOffset) return new CompileResult(eErrorType.Value);

            if(prevIsAddress && nextIsOffset)
            {
                return InternalCompile(Prev.Compile().Result.ToString(), Next.Compile().Result as ExcelDataProvider.IRangeInfo);
            }
            else if(prevIsOffset && nextIsAddress)
            {
                return InternalCompile(Prev.Compile().Result as ExcelDataProvider.IRangeInfo, Next.Compile().Result.ToString());
            }
            else if(prevIsOffset && nextIsOffset)
            {
                return InternalCompile(Prev.Compile().Result as ExcelDataProvider.IRangeInfo, Next.Compile().Result as ExcelDataProvider.IRangeInfo);
            }

            return new CompileResult(eErrorType.Value);
        }

        public override Expression MergeWithNext()
        {
            if(Prev.Prev != null)
            {
                Prev.Prev.Next = this;
            }
            if(Next.Next != null)
            {
                Next.Next.Prev = this;
            }
            return this;
        }

        private CompileResult InternalCompile(string address, ExcelDataProvider.IRangeInfo offsetRange)
        {
            throw new NotImplementedException();
        }

        private CompileResult InternalCompile(ExcelDataProvider.IRangeInfo offsetRange, string address)
        {
            throw new NotImplementedException();
        }

        private CompileResult InternalCompile(ExcelDataProvider.IRangeInfo offsetRange1, ExcelDataProvider.IRangeInfo offsetRange2)
        {
            throw new NotImplementedException();
        }
    }
}
