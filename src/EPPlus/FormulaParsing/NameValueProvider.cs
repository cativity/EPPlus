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

namespace OfficeOpenXml.FormulaParsing;

/// <summary>
/// Provides access to static, preconfigured instances of <see cref="INameValueProvider"/>
/// </summary>
public class NameValueProvider : INameValueProvider
{
    private NameValueProvider()
    {
    }

    /// <summary>
    /// An empty <see cref="INameValueProvider"/>
    /// </summary>
    public static INameValueProvider Empty => new NameValueProvider();

    /// <summary>
    /// Implementation of the IsNamedValue function. In this case (Empty provider) it always return false.
    /// </summary>
    /// <param name="key"></param>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public bool IsNamedValue(string key, string worksheet) => false;

    /// <summary>
    /// Implementation of the GetNamedValue function. In this case (Empty provider) it always return null.
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public object GetNamedValue(string key) => null;

    /// <summary>
    /// Implementation of the Reload function
    /// </summary>
    public void Reload()
    {
    }

    public object GetNamedValue(string key, string worksheet) => null;
}