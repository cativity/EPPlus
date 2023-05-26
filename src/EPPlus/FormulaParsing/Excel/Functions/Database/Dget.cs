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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Database,
        EPPlusVersion = "4",
        Description = "Returns a single value from a field of a list or database, that satisfy specified conditions")]
    internal class Dget : DatabaseFunction
    {

        public Dget()
            : this(new RowMatcher())
        {
            
        }

        public Dget(RowMatcher rowMatcher)
            : base(rowMatcher)
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            string? dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
            string? field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
            string? criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;

            ExcelDatabase? db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
            ExcelDatabaseCriteria? criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);

            int nHits = 0;
            object retVal = null;
            while (db.HasMoreRows)
            {
                ExcelDatabaseRow? dataRow = db.Read();
                if (!this.RowMatcher.IsMatch(dataRow, criteria))
                {
                    continue;
                }

                if(++nHits > 1)
                {
                    return this.CreateResult(ExcelErrorValue.Values.Num, DataType.ExcelError);
                }

                retVal = dataRow[field];
            }
            return new CompileResultFactory().Create(retVal);
        }
    }
}
