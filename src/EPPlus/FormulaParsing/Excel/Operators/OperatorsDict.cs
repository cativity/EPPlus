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

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public class OperatorsDict : Dictionary<string, IOperator>
    {
        public OperatorsDict()
        {
            this.Add("+", Operator.Plus);
            this.Add("-", Operator.Minus);
            this.Add("*", Operator.Multiply);
            this.Add("/", Operator.Divide);
            this.Add("^", Operator.Exp);
            this.Add("=", Operator.Eq);
            this.Add(">", Operator.GreaterThan);
            this.Add(">=", Operator.GreaterThanOrEqual);
            this.Add("<", Operator.LessThan);
            this.Add("<=", Operator.LessThanOrEqual);
            this.Add("<>", Operator.NotEqualsTo);
            this.Add("&", Operator.Concat);
        }

        private static IDictionary<string, IOperator> _instance;

        public static IDictionary<string, IOperator> Instance
        {
            get { return _instance ??= new OperatorsDict(); }
        }
    }
}
