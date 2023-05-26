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
            bool prevIsAddress = this.Prev.GetType() == typeof(ExcelAddressExpression);
            bool prevIsOffset = this.Prev.GetType() == typeof(FunctionExpression) && ((FunctionExpression)this.Prev).ExpressionString.ToLower() == "offset";
            bool nextIsAddress = this.Next.GetType() == typeof(ExcelAddressExpression);
            bool nextIsOffset = this.Next.GetType() == typeof(FunctionExpression) && ((FunctionExpression)this.Next).ExpressionString.ToLower() == "offset";

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
                return InternalCompile(this.Prev.Compile().Result.ToString(), this.Next.Compile().Result as IRangeInfo);
            }
            else if(prevIsOffset && nextIsAddress)
            {
                return InternalCompile(this.Prev.Compile().Result as IRangeInfo, this.Next.Compile().Result.ToString());
            }
            else if(prevIsOffset && nextIsOffset)
            {
                return InternalCompile(this.Prev.Compile().Result as IRangeInfo, this.Next.Compile().Result as IRangeInfo);
            }

            return new CompileResult(eErrorType.Value);
        }

        public override Expression MergeWithNext()
        {
            if(this.Prev.Prev != null)
            {
                this.Prev.Prev.Next = this;
            }
            if(this.Next.Next != null)
            {
                this.Next.Next.Prev = this;
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
