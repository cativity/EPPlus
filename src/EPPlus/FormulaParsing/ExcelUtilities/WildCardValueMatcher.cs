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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class WildCardValueMatcher : ValueMatcher
{
    protected override int CompareStringToString(string searchedValue, string candidate)
    {
        if (searchedValue.Contains("*") || searchedValue.Contains("?"))
        {
            string? regexPattern = BuildRegex(searchedValue, candidate);

            if (Regex.IsMatch(candidate, regexPattern))
            {
                return 0;
            }
        }

        return base.CompareStringToString(candidate, searchedValue);
    }

    private static string BuildRegex(string searchedValue, string candidate)
    {
        StringBuilder? result = new StringBuilder();
        string? regexPattern = Regex.Escape(searchedValue);
        regexPattern = regexPattern.Replace("\\*", "*");
        regexPattern = regexPattern.Replace("\\?", "?");
        regexPattern = string.Format("^{0}$", regexPattern);
        bool lastIsTilde = false;

        foreach (char ch in regexPattern)
        {
            if (ch == '~')
            {
                lastIsTilde = true;

                continue;
            }

            if (ch == '*')
            {
                if (lastIsTilde)
                {
                    _ = result.Append("\\*");
                }
                else
                {
                    _ = result.Append(".*");
                }
            }
            else if (ch == '?')
            {
                if (lastIsTilde)
                {
                    _ = result.Append("\\?");
                }
                else
                {
                    _ = result.Append('.');
                }
            }
            else if (lastIsTilde)
            {
                _ = result.Append("~" + ch);
            }
            else
            {
                _ = result.Append(ch);
            }

            lastIsTilde = false;
        }

        return result.ToString();
    }
}