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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;

/// <summary>
/// Postprocessor for a <see cref="TokenizerContext"/>
/// </summary>
internal class TokenizerPostProcessor
{
    public TokenizerPostProcessor(TokenizerContext context)
        : this(context, new TokenNavigator(context.Result))
    {
    }

    public TokenizerPostProcessor(TokenizerContext context, TokenNavigator navigator)
    {
        this._context = context;
        this._navigator = navigator;
    }

    private readonly TokenizerContext _context;
    private readonly TokenNavigator _navigator;
    private readonly Token PlusToken = TokenSeparatorProvider.Instance.GetToken("+").Value;
    private readonly Token MinusToken = TokenSeparatorProvider.Instance.GetToken("-").Value;

    /// <summary>
    /// Processes the <see cref="TokenizerContext"/>
    /// </summary>
    public void Process()
    {
        bool hasColon = false;

        while (this._navigator.HasNext())
        {
            Token token = this._navigator.CurrentToken;

            if (token.TokenTypeIsSet(TokenType.Unrecognized))
            {
                this.HandleUnrecognizedToken();
            }
            else if (token.TokenTypeIsSet(TokenType.Colon))
            {
                this.HandleColon();
                hasColon = true;
            }
            else if (token.TokenTypeIsSet(TokenType.WorksheetName))
            {
                this.HandleWorksheetNameToken();
            }
            else if (token.TokenTypeIsSet(TokenType.Operator) || token.TokenTypeIsSet(TokenType.Negator))
            {
                if (token.Value == "+" || token.Value == "-")
                {
                    this.HandleNegators();
                }
            }

            this._navigator.MoveNext();
        }

        if (hasColon)
        {
            this._navigator.MoveIndex(-this._navigator.Index);

            while (this._navigator.HasNext())
            {
                Token token = this._navigator.CurrentToken;

                if (token.TokenTypeIsSet(TokenType.Colon) && this._context.Result.Count > this._navigator.Index + 1)
                {
                    if (this._navigator.PreviousToken != null
                        && this._navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.ExcelAddress)
                        && this._navigator.NextToken.TokenTypeIsSet(TokenType.ExcelAddress))
                    {
                        string? newToken = this._navigator.PreviousToken.Value.Value + ":" + this._navigator.NextToken.Value;
                        this._context.Result[this._navigator.Index - 1] = new Token(newToken, TokenType.ExcelAddress);
                        this._context.RemoveAt(this._navigator.Index);
                        this._context.RemoveAt(this._navigator.Index);
                        this._navigator.MoveIndex(-1);
                    }
                }

                this._navigator.MoveNext();
            }
        }
    }

    private void ChangeTokenTypeOnCurrentToken(TokenType tokenType)
    {
        this._context.ChangeTokenType(tokenType, this._navigator.Index);
    }

    private void ChangeValueOnCurrentToken(string value)
    {
        this._context.ChangeValue(value, this._navigator.Index);
    }

    private static bool IsOffsetFunctionToken(Token token)
    {
        return token.TokenTypeIsSet(TokenType.Function) && token.Value.ToLower() == "offset";
    }

    private void HandleColon()
    {
        Token prevToken = this._navigator.GetTokenAtRelativePosition(-1);
        Token nextToken = this._navigator.GetTokenAtRelativePosition(1);

        if (prevToken.TokenTypeIsSet(TokenType.ClosingParenthesis))
        {
            // Previous expression should be an OFFSET function
            int index = 0;
            int openedParenthesis = 0;
            int closedParethesis = 0;

            while (openedParenthesis == 0 || openedParenthesis > closedParethesis)
            {
                index--;
                Token token = this._navigator.GetTokenAtRelativePosition(index);

                if (token.TokenTypeIsSet(TokenType.ClosingParenthesis))
                {
                    openedParenthesis++;
                }
                else if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    closedParethesis++;
                }
            }

            Token offsetCandidate = this._navigator.GetTokenAtRelativePosition(--index);

            if (IsOffsetFunctionToken(offsetCandidate))
            {
                this._context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, this._navigator.Index + index);

                if (nextToken.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    // OFFSET:A1
                    this._context.ChangeTokenType(TokenType.ExcelAddress | TokenType.RangeOffset, this._navigator.Index + 1);
                }
                else if (IsOffsetFunctionToken(nextToken))
                {
                    // OFFSET:OFFSET
                    this._context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, this._navigator.Index + 1);
                }
            }
        }
        else if (prevToken.TokenTypeIsSet(TokenType.ExcelAddress) && IsOffsetFunctionToken(nextToken))
        {
            // A1: OFFSET
            this._context.ChangeTokenType(TokenType.ExcelAddress | TokenType.RangeOffset, this._navigator.Index - 1);
            this._context.ChangeTokenType(TokenType.Function | TokenType.RangeOffset, this._navigator.Index + 1);
        }
    }

    private void HandleNegators()
    {
        Token token = this._navigator.CurrentToken;

        //Remove '+' from start of formula and formula arguments
        if (token.Value == "+"
            && (!this._navigator.HasPrev()
                || this._navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.OpeningParenthesis)
                || this._navigator.PreviousToken.Value.TokenTypeIsSet(TokenType.Comma)))
        {
            this.RemoveTokenAndSetNegatorOperator();

            return;
        }

        Token nextToken = this._navigator.NextToken;

        if (nextToken.TokenTypeIsSet(TokenType.Operator) || nextToken.TokenTypeIsSet(TokenType.Negator))
        {
            // Remove leading '+' from operator combinations
            if (token.Value == "+" && (nextToken.Value == "+" || nextToken.Value == "-"))
            {
                this.RemoveTokenAndSetNegatorOperator();
            }

            // Remove trailing '+' from a negator operation
            else if (token.Value == "-" && nextToken.Value == "+")
            {
                this.RemoveTokenAndSetNegatorOperator(1);
            }

            // Convert double negator operation to positive declaration
            else if (token.Value == "-" && nextToken.Value == "-")
            {
                this._context.ChangeTokenType(TokenType.Negator, this._navigator.Index);
                this._navigator.MoveIndex(1);
                this._context.ChangeTokenType(TokenType.Negator, this._navigator.Index);
                /*
                _context.RemoveAt(_navigator.Index);
                _context.Replace(_navigator.Index, PlusToken);
                if (_navigator.Index > 0) _navigator.MoveIndex(-1);
                HandleNegators();
                */
            }
        }
    }

    private void HandleUnrecognizedToken()
    {
        if (this._navigator.HasNext())
        {
            if (this._navigator.NextToken.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                this.ChangeTokenTypeOnCurrentToken(TokenType.Function);
            }
            else
            {
                this.ChangeTokenTypeOnCurrentToken(TokenType.NameValue);
            }
        }
        else
        {
            this.ChangeTokenTypeOnCurrentToken(TokenType.NameValue);
        }
    }

    private void HandleWorksheetNameToken()
    {
        // use this and the following three tokens
        Token relativeToken = this._navigator.GetTokenAtRelativePosition(3);
        TokenType tokenType = relativeToken.GetTokenTypeFlags();
        this.ChangeTokenTypeOnCurrentToken(tokenType);
        StringBuilder? sb = new StringBuilder();
        int nToRemove = 3;

        if (this._navigator.NbrOfRemainingTokens < nToRemove)
        {
            this.ChangeTokenTypeOnCurrentToken(TokenType.InvalidReference);
            nToRemove = this._navigator.NbrOfRemainingTokens;
        }

        if (relativeToken.TokenTypeIsSet(TokenType.Comma) || relativeToken.TokenTypeIsSet(TokenType.ClosingParenthesis))
        {
            for (int ix = 0; ix < 3; ix++)
            {
                _ = sb.Append(this._navigator.GetTokenAtRelativePosition(ix).Value);
            }

            this.ChangeTokenTypeOnCurrentToken(TokenType.ExcelAddress);
            nToRemove = 2;
        }
        else if (!relativeToken.TokenTypeIsSet(TokenType.ExcelAddress)
                 && !relativeToken.TokenTypeIsSet(TokenType.ExcelAddressR1C1)
                 && !relativeToken.TokenTypeIsSet(TokenType.NameValue)
                 && !relativeToken.TokenTypeIsSet(TokenType.InvalidReference))
        {
            this.ChangeTokenTypeOnCurrentToken(TokenType.InvalidReference);
            nToRemove--;
        }
        else
        {
            for (int ix = 0; ix < 4; ix++)
            {
                _ = sb.Append(this._navigator.GetTokenAtRelativePosition(ix).Value);
            }
        }

        this.ChangeValueOnCurrentToken(sb.ToString());

        for (int ix = 0; ix < nToRemove; ix++)
        {
            this._context.RemoveAt(this._navigator.Index + 1);
        }
    }

    private void SetNegatorOperator(int i)
    {
        Token token = this._context.Result[i];

        if (token.Value == "-" && i > 0 && (token.TokenTypeIsSet(TokenType.Operator) || token.TokenTypeIsSet(TokenType.Negator)))
        {
            if (TokenIsNegator(this._context.Result[i - 1]))
            {
                this._context.Replace(i, new Token("-", TokenType.Negator));
            }
            else
            {
                this._context.Replace(i, this.MinusToken);
            }
        }
    }

    private static bool TokenIsNegator(TokenizerContext context)
    {
        return TokenIsNegator(context.LastToken.Value);
    }

    private static bool TokenIsNegator(Token t)
    {
        return t.TokenTypeIsSet(TokenType.Operator)
               || t.TokenTypeIsSet(TokenType.OpeningParenthesis)
               || t.TokenTypeIsSet(TokenType.Comma)
               || t.TokenTypeIsSet(TokenType.SemiColon)
               || t.TokenTypeIsSet(TokenType.OpeningEnumerable);
    }

    private void RemoveTokenAndSetNegatorOperator(int offset = 0)
    {
        this._context.Result.RemoveAt(this._navigator.Index + offset);
        this.SetNegatorOperator(this._navigator.Index);

        if (this._navigator.Index > 0)
        {
            this._navigator.MoveIndex(-1);
        }
    }
}