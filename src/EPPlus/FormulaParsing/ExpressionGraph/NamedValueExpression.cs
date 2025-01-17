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
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Table;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class NamedValueExpression : AtomicExpression
{
    public NamedValueExpression(string expression, ParsingContext parsingContext)
        : base(expression) =>
        this._parsingContext = parsingContext;

    private readonly ParsingContext _parsingContext;

    public override CompileResult Compile()
    {
        ParsingScope? c = this._parsingContext.Scopes.Current;
        INameInfo? name = this._parsingContext.ExcelDataProvider.GetName(c.Address.Worksheet, this.ExpressionString);

        ExcelAddressCache? cache = this._parsingContext.AddressCache;
        int cacheId = cache.GetNewId();

        if (name == null)
        {
            // check if there is a table with the name
            ExcelTable? table = this._parsingContext.ExcelDataProvider.GetExcelTable(this.ExpressionString);

            if (table != null)
            {
                RangeInfo? ri = new RangeInfo(table.WorkSheet, table.Address);
                _ = cache.Add(cacheId, ri.Address.FullAddress);

                return new CompileResult(ri, DataType.Enumerable, cacheId);
            }

            return new CompileResult(eErrorType.Name);
        }

        if (name.Value == null)
        {
            return new CompileResult(null, DataType.Empty, cacheId);
        }

        if (name.Value is IRangeInfo)
        {
            IRangeInfo? range = (IRangeInfo)name.Value;
            _ = cache.Add(cacheId, range.Address.FullAddress);

            if (range.IsMulti)
            {
                return new CompileResult(name.Value, DataType.Enumerable, cacheId);
            }
            else
            {
                if (range.IsEmpty)
                {
                    return new CompileResult(null, DataType.Empty, cacheId);
                }

                CompileResultFactory? factory = new CompileResultFactory();

                return factory.Create(range.First().Value, cacheId);
            }
        }
        else
        {
            CompileResultFactory? factory = new CompileResultFactory();

            return factory.Create(name.Value, cacheId);
        }

        //return new CompileResultFactory().Create(result);
    }
}