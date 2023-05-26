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
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;

/// <summary>
/// Helper class for reading and modifying tokens compiled by the <see cref="TokenizerContext"/>
/// </summary>
public class TokenNavigator
{
    public TokenNavigator(IList<Token> tokens)
    {
        this._tokens = tokens;
    }

    private readonly IList<Token> _tokens;

    /// <summary>
    /// Returns true if there is a next token relative to the current token.
    /// </summary>
    /// <returns></returns>
    public bool HasNext()
    {
        return this.Index < this._tokens.Count - 1;
    }

    /// <summary>
    /// Returns true if there is a previous token relative to the current token.
    /// </summary>
    /// <returns></returns>
    public bool HasPrev()
    {
        return this.Index > 0;
    }

    /// <summary>
    /// Moves to the next token
    /// </summary>
    public void MoveNext()
    {
        this.Index++;
    }

    /// <summary>
    /// The index of the current token.
    /// </summary>
    public int Index { get; private set; } = 0;

    /// <summary>
    /// Remaining number of tokens
    /// </summary>
    public int NbrOfRemainingTokens
    {
        get { return this._tokens.Count - this.Index - 1; }
    }

    /// <summary>
    /// The current token.
    /// </summary>
    public Token CurrentToken
    {
        get { return this._tokens[this.Index]; }
    }

    /// <summary>
    /// The token before the current token. If current token is the first token, null will be returned.
    /// </summary>
    public Token? PreviousToken
    {
        get { return this.Index == 0 ? default(Token?) : this._tokens[this.Index - 1]; }
    }

    public Token NextToken
    {
        get
        {
            if (this.Index == this._tokens.Count - 1)
            {
                throw new ArgumentOutOfRangeException("NextToken: current token is the last token");
            }

            return this._tokens[this.Index + 1];
        }
    }
        
    /// <summary>
    /// Moves to a position relative to current token
    /// </summary>
    /// <param name="relativePosition">The requested position relative to current</param>
    public void MoveIndex(int relativePosition)
    {
        int newPosition = this.Index + relativePosition;
        if (newPosition < 0 || newPosition >= this._tokens.Count)
        {
            throw new ArgumentOutOfRangeException("MoveIndex: new index out of range");
        }

        this.Index += relativePosition;
    }

    /// <summary>
    /// Returns a token using a relative position (offset) of the current token.
    /// </summary>
    /// <param name="relativePosition">Offset, can be positive or negative</param>
    /// <returns>The <see cref="Token"/> of the requested position</returns>
    public Token GetTokenAtRelativePosition(int relativePosition)
    {
        int newPosition = this.Index + relativePosition;
        if (newPosition < 0 || newPosition >= this._tokens.Count)
        {
            throw new ArgumentOutOfRangeException("¨GetTokenAtRelativePosition: new index out of range");
        }

        return this._tokens[newPosition];
    }
}