using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class ColonExpression : FunctionExpression
    {
        public ColonExpression(string expression, ParsingContext parsingContext) : base(expression, parsingContext, false)
        {
        }

        public override CompileResult Compile()
        {
            bool prevIsAddress = Prev.GetType() == typeof(ExcelAddressExpression);
            bool prevIsOffset = Prev.GetType() == typeof(FunctionExpression) && ((FunctionExpression)Prev).ExpressionString.ToLower() == "offset";
            bool nextIsAddress = Next.GetType() == typeof(ExcelAddressExpression);
            bool nextIsOffset = Next.GetType() == typeof(FunctionExpression) && ((FunctionExpression)Next).ExpressionString.ToLower() == "offset";

            if (!prevIsAddress && !prevIsOffset)
            {
                return new CompileResult(eErrorType.Value);
            }

            if (!nextIsAddress && !nextIsOffset)
            {
                return new CompileResult(eErrorType.Value);
            }

            if(prevIsAddress && nextIsOffset)
            {
                return InternalCompile(Prev.Compile().Result.ToString(), Next.Compile().Result as IRangeInfo);
            }
            else if(prevIsOffset && nextIsAddress)
            {
                return InternalCompile(Prev.Compile().Result as IRangeInfo, Next.Compile().Result.ToString());
            }
            else if(prevIsOffset && nextIsOffset)
            {
                return InternalCompile(Prev.Compile().Result as IRangeInfo, Next.Compile().Result as IRangeInfo);
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

        private static CompileResult InternalCompile(string address, IRangeInfo offsetRange)
        {
            throw new NotImplementedException();
        }

        private static CompileResult InternalCompile(IRangeInfo offsetRange, string address)
        {
            throw new NotImplementedException();
        }

        private static CompileResult InternalCompile(IRangeInfo offsetRange1, IRangeInfo offsetRange2)
        {
            throw new NotImplementedException();
        }
    }
}
