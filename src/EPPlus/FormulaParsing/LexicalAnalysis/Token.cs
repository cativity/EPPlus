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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

/// <summary>
/// Represents a character in a formula
/// </summary>
public struct Token
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="token">The formula character</param>
    /// <param name="tokenType">The <see cref="TokenType"/></param>
    public Token(string token, TokenType tokenType)
        : this(token, tokenType, false)
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="token">The formula character</param>
    /// <param name="tokenType">The <see cref="TokenType"></see></param>
    /// <param name="isNegated"></param>
    public Token(string token, TokenType tokenType, bool isNegated)
    {
        this.Value = token;
        this._tokenType = tokenType;
        this.IsNegated = isNegated;
    }

    /// <summary>
    /// The formula character
    /// </summary>
    public string Value { get; internal set; }

    private TokenType _tokenType;

    /// <summary>
    /// Indicates whether a numeric value should be negated when compiled
    /// </summary>
    public bool IsNegated { get; private set; }

    /// <summary>
    /// Operator ==
    /// </summary>
    /// <param name="t1"></param>
    /// <param name="t2"></param>
    /// <returns></returns>
    public static bool operator ==(Token t1, Token t2) => t1.AreEqualTo(t2);

    /// <summary>
    /// Operator !=
    /// </summary>
    /// <param name="t1"></param>
    /// <param name="t2"></param>
    /// <returns></returns>
    public static bool operator !=(Token t1, Token t2) => !t1.AreEqualTo(t2);

    /// <summary>
    /// Overrides object.Equals with no behavioural change
    /// </summary>
    /// <param name="obj"></param>
    /// <returns></returns>
    public override bool Equals(object obj) => base.Equals(obj);

    /// <summary>
    /// Overrides object.GetHashCode with no behavioural change
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => base.GetHashCode();

    /// <summary>
    /// Return if the supplied <paramref name="tokenType"/> is set on this token.
    /// </summary>
    /// <param name="tokenType">The <see cref="TokenType"></see> to check</param>
    /// <returns>True if the token is set, otherwirse false</returns>
    public bool TokenTypeIsSet(TokenType tokenType) => (this._tokenType & tokenType) == tokenType;

    public bool AreEqualTo(Token otherToken) => this.GetTokenTypeFlags() == otherToken.GetTokenTypeFlags() && this.Value == otherToken.Value;

    internal TokenType GetTokenTypeFlags() => this._tokenType;

    /// <summary>
    /// Clones the token with a new <see cref="TokenType"/> set.
    /// </summary>
    /// <param name="tokenType">The new TokenType</param>
    /// <returns>A cloned Token</returns>
    internal Token CloneWithNewTokenType(TokenType tokenType) => new(this.Value, tokenType, this.IsNegated);

    /// <summary>
    /// Clones the token with a new value set.
    /// </summary>
    /// <param name="val">The new value</param>
    /// <returns>A cloned Token</returns>
    internal Token CloneWithNewValue(string val) => new(val, this._tokenType, this.IsNegated);

    /// <summary>
    /// Clones the token with a new value set for isNegated.
    /// </summary>
    /// <param name="isNegated">The new isNegated value</param>
    /// <returns>A cloned Token</returns>
    internal Token CloneWithNegatedValue(bool isNegated)
    {
        if ((this._tokenType & TokenType.Decimal) == 0 || (this._tokenType & TokenType.Integer) == 0 || (this._tokenType & TokenType.ExcelAddress) == 0)
        {
            return new Token(this.Value, this._tokenType, isNegated);
        }

        return this;
    }

    /// <summary>
    /// Overrides object.ToString()
    /// </summary>
    /// <returns>TokenType, followed by value</returns>
    public override string ToString() => this._tokenType.ToString() + ", " + this.Value;
}