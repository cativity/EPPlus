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
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing;

internal class EpplusNameValueProvider : INameValueProvider
{
    private ExcelDataProvider _excelDataProvider;
    private ExcelNamedRangeCollection _values;

    internal EpplusNameValueProvider(ExcelDataProvider excelDataProvider)
    {
        this._excelDataProvider = excelDataProvider;
        this._values = this._excelDataProvider.GetWorkbookNameValues();
    }

    public virtual bool IsNamedValue(string key, string ws)
    {
        if (key.StartsWith("[0]"))
        {
            if(key.Length>3&&key[3]=='!')
            {
                key = key.Substring(4);
            }
            else
            {
                key = key.Substring(3);
            }
        }
        if (key.StartsWith("["))
        {
            return this._excelDataProvider.IsExternalName(key);
        }
        else if (ws!=null)
        {
            ExcelNamedRangeCollection? wsNames = this._excelDataProvider.GetWorksheetNames(ws);
            if (wsNames != null && wsNames.ContainsKey(key))
            {
                return true;
            }
        }
        return this._values != null && this._values.ContainsKey(key);
    }

    public virtual object GetNamedValue(string key)
    {
        return this._values[key];
    }

    public virtual object GetNamedValue(string key, string worksheet)
    {
        return this._excelDataProvider.GetName(worksheet, key)?.Value;
    }

    public virtual void Reload()
    {
        this._values = this._excelDataProvider.GetWorkbookNameValues();
    }
}