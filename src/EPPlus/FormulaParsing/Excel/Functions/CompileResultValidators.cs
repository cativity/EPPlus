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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class CompileResultValidators
{
    private readonly Dictionary<DataType, CompileResultValidator> _validators = new Dictionary<DataType, CompileResultValidator>();

    private CompileResultValidator CreateOrGet(DataType dataType)
    {
        if (this._validators.ContainsKey(dataType))
        {
            return this._validators[dataType];
        }

        if (dataType == DataType.Decimal)
        {
            return this._validators[DataType.Decimal] = new DecimalCompileResultValidator();
        }

        return CompileResultValidator.Empty;
    }

    public CompileResultValidator GetValidator(DataType dataType) => this.CreateOrGet(dataType);
}