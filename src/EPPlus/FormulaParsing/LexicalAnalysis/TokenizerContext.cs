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

using OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

/// <summary>
/// Responsible for handling tokens during the tokenizing process.
/// </summary>
internal class TokenizerContext
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="formula">The formula to tokenize</param>
    /// <param name="worksheet">Worksheet name, if applicable</param>
    /// <param name="tokenFactory">A <see cref="ITokenFactory"/> instance</param>
    public TokenizerContext(string formula, string worksheet, ITokenFactory tokenFactory)
    {
        if (!string.IsNullOrEmpty(formula))
        {
            this.FormulaChars = formula.ToArray();
        }

        Require.That(tokenFactory).IsNotNull();
        this._result = new List<Token>();
        this._currentToken = new StringBuilder();
        this._worksheet = worksheet;
        this._tokenFactory = tokenFactory;
    }

    private readonly List<Token> _result;
    private StringBuilder _currentToken;
    private readonly ITokenFactory _tokenFactory;
    private readonly string _worksheet;

    /// <summary>
    /// The formula split into a character array
    /// </summary>
    public char[] FormulaChars { get; private set; }

    public TokenHandler CreateHandler(INameValueProvider nameValueProvider)
    {
        TokenHandler? handler = new TokenHandler(this, this._tokenFactory, TokenSeparatorProvider.Instance, nameValueProvider);
        handler.Worksheet = this._worksheet;

        return handler;
    }

    /// <summary>
    /// The tokens created
    /// </summary>
    public IList<Token> Result => this._result;

    internal string Worksheet => this._worksheet;

    /// <summary>
    /// Returns the token before the requested index
    /// </summary>
    /// <param name="index">The requested index</param>
    /// <returns>The <see cref="Token"/> at the requested position</returns>
    public Token GetTokenBeforeIndex(int index)
    {
        if (index < 1 || index > this._result.Count - 1)
        {
            throw new IndexOutOfRangeException("Index was out of range of the token array");
        }

        return this._result[index - 1];
    }

    /// <summary>
    /// Returns the token after the requested index
    /// </summary>
    /// <param name="index">The requested index</param>
    /// <returns>The <see cref="Token"/> at the requested position</returns>
    public Token GetNextTokenAfterIndex(int index)
    {
        if (index < 0 || index > this._result.Count - 2)
        {
            throw new IndexOutOfRangeException("Index was out of range of the token array");
        }

        return this._result[index + 1];
    }

    private Token CreateToken(string worksheet)
    {
        if (this.CurrentToken == "-")
        {
            if (this.LastToken == null && this.LastToken.Value.TokenTypeIsSet(TokenType.Operator))
            {
                return new Token("-", TokenType.Negator);
            }
        }

        return this._tokenFactory.Create(this.Result, this.CurrentToken, worksheet);
    }

    internal void OverwriteCurrentToken(string token) => this._currentToken = new StringBuilder(token);

    public void PostProcess()
    {
        if (this.CurrentTokenHasValue)
        {
            this.AddToken(this.CreateToken(this._worksheet));
        }

        TokenizerPostProcessor? postProcessor = new TokenizerPostProcessor(this);
        postProcessor.Process();
    }

    /// <summary>
    /// Replaces a token at the requested <paramref name="index"/>
    /// </summary>
    /// <param name="index">0-based index of the requested position</param>
    /// <param name="newValue">The new <see cref="Token"/></param>
    public void Replace(int index, Token newValue) => this._result[index] = newValue;

    /// <summary>
    /// Removes the token at the requested <see cref="Token"/>
    /// </summary>
    /// <param name="index">0-based index of the requested position</param>
    public void RemoveAt(int index) => this._result.RemoveAt(index);

    /// <summary>
    /// Returns true if the current position is inside a string, otherwise false.
    /// </summary>
    public bool IsInString { get; private set; }

    /// <summary>
    /// Returns true if the current position is inside a sheetname, otherwise false.
    /// </summary>
    public bool IsInSheetName { get; private set; }

    /// <summary>
    /// Toggles the IsInString state.
    /// </summary>
    public void ToggleIsInString() => this.IsInString = !this.IsInString;

    /// <summary>
    /// Toggles the IsInSheetName state
    /// </summary>
    public void ToggleIsInSheetName() => this.IsInSheetName = !this.IsInSheetName;

    internal int BracketCount { get; set; }

    internal bool IsInDefinedNameAddress { get; set; }

    /// <summary>
    /// Returns the current
    /// </summary>
    public string CurrentToken => this._currentToken.ToString();

    public bool CurrentTokenHasValue => !string.IsNullOrEmpty(this.IsInString ? this.CurrentToken : this.CurrentToken.Trim());

    public void NewToken() => this._currentToken = new StringBuilder();

    public void AddToken(Token token) => this._result.Add(token);

    public void AppendToCurrentToken(char c) => _ = this._currentToken.Append(c.ToString());

    public void AppendToLastToken(string stringToAppend)
    {
        Token token = this._result.Last();
        string? newVal = token.Value += stringToAppend;
        Token newToken = token.CloneWithNewValue(newVal);
        this.ReplaceLastToken(newToken);
    }

    /// <summary>
    /// Changes <see cref="TokenType"/> of the current token.
    /// </summary>
    /// <param name="tokenType">The new <see cref="TokenType"/></param>
    /// <param name="index">Index of the token to change</param>
    public void ChangeTokenType(TokenType tokenType, int index) => this._result[index] = this._result[index].CloneWithNewTokenType(tokenType);

    /// <summary>
    /// Changes the value of the current token
    /// </summary>
    /// <param name="val"></param>
    /// <param name="index">Index of the token to change</param>
    public void ChangeValue(string val, int index) => this._result[index] = this._result[index].CloneWithNewValue(val);

    /// <summary>
    /// Changes the <see cref="TokenType"/> of the last token in the result.
    /// </summary>
    /// <param name="type"></param>
    public void SetLastTokenType(TokenType type)
    {
        Token newToken = this._result.Last().CloneWithNewTokenType(type);
        this.ReplaceLastToken(newToken);
    }

    /// <summary>
    /// Replaces the last token of the result with the <paramref name="newToken"/>
    /// </summary>
    /// <param name="newToken">The new token</param>
    public void ReplaceLastToken(Token newToken)
    {
        int count = this._result.Count;

        if (count > 0)
        {
            this._result.RemoveAt(count - 1);
        }

        this._result.Add(newToken);
    }

    /// <summary>
    /// Returns the last token of the result, if empty null/default(Token?) will be returned.
    /// </summary>
    public Token? LastToken => this._result.Count > 0 ? this._result.Last() : default(Token?);
}