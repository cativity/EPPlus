/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/04/2022         EPPlus Software AB       Initial release EPPlus 6.1
 *************************************************************************************************/

using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace OfficeOpenXml.Export.ToCollection;

internal class ToCollectionAutomap
{
    internal static List<MappedProperty> GetAutomapList<T>(List<string> h)
    {
        Type? t = typeof(T);

        List<MappedProperty>? pl = new List<MappedProperty>();

        foreach (PropertyInfo? m in t.GetProperties(BindingFlags.Instance | BindingFlags.Public))
        {
            int ix = h.FindIndex(x => RemoveWS(x).Equals(m.Name, StringComparison.CurrentCultureIgnoreCase));

            if (ix < 0)
            {
                EpplusTableColumnAttributeBase? tca = m.GetFirstAttributeOfType<EpplusTableColumnAttributeBase>(true);

                if (tca != null)
                {
                    ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(tca.Header), StringComparison.CurrentCultureIgnoreCase));
                }

                if (ix < 0)
                {
                    DescriptionAttribute? da = m.GetFirstAttributeOfType<DescriptionAttribute>();

                    if (da != null)
                    {
                        ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(da.Description), StringComparison.CurrentCultureIgnoreCase));
                    }

                    if (ix < 0)
                    {
                        DisplayNameAttribute? dna = m.GetFirstAttributeOfType<DisplayNameAttribute>();

                        if (dna != null)
                        {
                            ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(dna.DisplayName), StringComparison.CurrentCultureIgnoreCase));
                        }
                    }
                }
            }

            if (ix >= 0)
            {
                pl.Add(new MappedProperty(ix, m));
            }
        }

        return pl;
    }

    private static string RemoveWS(string v)
    {
        if (v != null)
        {
            return v.Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
        }

        return v;
    }
}