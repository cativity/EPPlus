/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/16/2022         EPPlus Software AB       Fix for issue #593
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;

internal class DefinedNameAddressHandler : SeparatorHandler
{
    INameValueProvider _nameValueProvider;

    public DefinedNameAddressHandler(INameValueProvider nameValueProvider) => this._nameValueProvider = nameValueProvider;

    public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
    {
        if (context.IsInDefinedNameAddress && (c == ')' || c == ','))
        {
            if (context.IsInDefinedNameAddress)
            {
                context.IsInDefinedNameAddress = false;

                // the first name is already resolved to an address followed by a dot
                string? tokenValue = context.CurrentToken?.ToString();

                if (!string.IsNullOrEmpty(tokenValue))
                {
                    string[]? parts = tokenValue.Split(':');

                    if (parts.Length < 2)
                    {
                        return false;
                    }

                    string? part1 = parts[0];
                    string? name = parts[1];
                    object? nameValue = this._nameValueProvider.GetNamedValue(name, context.Worksheet);

                    if (nameValue != null)
                    {
                        if (nameValue is IRangeInfo rangeInfo)
                        {
                            string? address = part1 + ":" + rangeInfo.Address.Address;
                            Token addressToken = new Token(address, TokenType.ExcelAddress);
                            context.AddToken(addressToken);
                            context.AddToken(tokenSeparator);
                            context.NewToken();

                            return true;
                        }
                    }
                }
            }
        }

        return false;
    }
}