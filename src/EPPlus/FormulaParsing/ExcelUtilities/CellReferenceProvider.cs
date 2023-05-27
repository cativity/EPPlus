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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class CellReferenceProvider
{
    public virtual IEnumerable<string> GetReferencedAddresses(string cellFormula, ParsingContext context)
    {
        List<string>? resultCells = new List<string>();
        IEnumerable<Token>? r = context.Configuration.Lexer.Tokenize(cellFormula, context.Scopes.Current.Address.Worksheet);
        IEnumerable<Token>? toAddresses = r.Where(x => x.TokenTypeIsSet(TokenType.ExcelAddress));

        foreach (Token toAddress in toAddresses)
        {
            RangeAddress? rangeAddress = context.RangeAddressFactory.Create(toAddress.Value);
            List<string>? rangeCells = new List<string>();

            if (rangeAddress.FromRow < rangeAddress.ToRow || rangeAddress.FromCol < rangeAddress.ToCol)
            {
                for (int col = rangeAddress.FromCol; col <= rangeAddress.ToCol; col++)
                {
                    for (int row = rangeAddress.FromRow; row <= rangeAddress.ToRow; row++)
                    {
                        resultCells.Add(context.RangeAddressFactory.Create(col, row).Address);
                    }
                }
            }
            else
            {
                rangeCells.Add(toAddress.Value);
            }

            resultCells.AddRange(rangeCells);
        }

        return resultCells;
    }
}