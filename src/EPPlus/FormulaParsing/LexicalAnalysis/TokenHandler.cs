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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

internal class TokenHandler : ITokenIndexProvider
{
    public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider, INameValueProvider nameValueProvider)
        : this(context, tokenFactory, tokenProvider, new TokenSeparatorHandler(tokenProvider, nameValueProvider))
    {

    }
    public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider, TokenSeparatorHandler tokenSeparatorHandler)
    {
        this._context = context;
        this._tokenFactory = tokenFactory;
        this._tokenProvider = tokenProvider;
        this._tokenSeparatorHandler = tokenSeparatorHandler;
    }

    private readonly TokenizerContext _context;
    private readonly ITokenSeparatorProvider _tokenProvider;
    private readonly ITokenFactory _tokenFactory;
    private readonly TokenSeparatorHandler _tokenSeparatorHandler;
    private int _tokenIndex = -1;

    public string Worksheet { get; set; }

    public bool HasMore()
    {
        return this._tokenIndex < this._context.FormulaChars.Length - 1;
    }

    public void Next()
    {
        this._tokenIndex++;
        this.Handle();
    }

    private void Handle()
    {
        char c = this._context.FormulaChars[this._tokenIndex];

        if (this.CharIsTokenSeparator(c, out Token tokenSeparator))
        {
            if (this._tokenSeparatorHandler.Handle(c, tokenSeparator, this._context, this))
            {
                return;
            }
                                              
            if (this._context.CurrentTokenHasValue)
            {
                //if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
                if(this._context.CurrentToken.StartsWith("\"") && this._context.CurrentToken.EndsWith("\""))
                {
                    this._context.AddToken(this._tokenFactory.Create(this._context.CurrentToken, TokenType.StringContent));
                }
                else
                {
                    this._context.AddToken(this.CreateToken(this._context, this.Worksheet));
                }


                //If the a next token is an opening parantheses and the previous token is interpeted as an address or name, then the currenct token is a function
                if (tokenSeparator.TokenTypeIsSet(TokenType.OpeningParenthesis) && (this._context.LastToken.Value.TokenTypeIsSet(TokenType.ExcelAddress) || this._context.LastToken.Value.TokenTypeIsSet(TokenType.NameValue)))
                {
                    Token newToken = this._context.LastToken.Value.CloneWithNewTokenType(TokenType.Function);
                    this._context.ReplaceLastToken(newToken);
                }
            }
            if (TokenIsNegator(tokenSeparator.Value, this._context))
            {
                this._context.AddToken(new Token("-", TokenType.Negator));
                return;
            }

            this._context.AddToken(tokenSeparator);
            this._context.NewToken();
            return;
        }

        this._context.AppendToCurrentToken(c);
    }

    private bool CharIsTokenSeparator(char c, out Token token)
    {
        bool result = this._tokenProvider.Tokens.ContainsKey(c.ToString());
        token = result ? token = this._tokenProvider.Tokens[c.ToString()] : default(Token);
        return result;
    }

    private static bool TokenIsNegator(string token, TokenizerContext context)
    {
        if (token != "-")
        {
            return false;
        }

        if (!context.LastToken.HasValue)
        {
            return true;
        }

        Token t = context.LastToken.Value;
            
        return t.TokenTypeIsSet(TokenType.Operator)
               ||
               t.TokenTypeIsSet(TokenType.OpeningParenthesis)
               ||
               t.TokenTypeIsSet(TokenType.Comma)
               ||
               t.TokenTypeIsSet(TokenType.SemiColon)
               ||
               t.TokenTypeIsSet(TokenType.OpeningEnumerable);
    }

    private Token CreateToken(TokenizerContext context, string worksheet)
    {
        if (context.CurrentToken == "-")
        {
            if (context.LastToken == default(Token) && context.LastToken.Value.TokenTypeIsSet(TokenType.Operator))
            {
                return new Token("-", TokenType.Negator);
            }
        }
        return this._tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
    }

    int ITokenIndexProvider.Index
    {
        get { return this._tokenIndex; }
    }


    void ITokenIndexProvider.MoveIndexPointerForward()
    {
        this._tokenIndex++;
    }
}