﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  02/03/2020         EPPlus Software AB       Added
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml.Core.Worksheet;

internal class FormulaDataTableValidation
{
    internal static void HasPartlyFormulaDataTable(ExcelWorksheet ws, ExcelAddressBase deleteRange, bool isDelete, string errorMsg)
    {
        HashSet<int>? hs = new HashSet<int>();

        CellStoreEnumerator<object>? cse =
            new CellStoreEnumerator<object>(ws._formulas, deleteRange._fromRow, deleteRange._fromCol, deleteRange._toRow + 1, deleteRange._toCol + 1);

        while (cse.Next())
        {
            if (cse.Value is int si && hs.Contains(si) == false)
            {
                ExcelWorksheet.Formulas? f = ws._sharedFormulas[si];

                if (f.FormulaType == ExcelWorksheet.FormulaType.DataTable)
                {
                    ExcelAddressBase? fa = new ExcelAddressBase(f.Address);

                    if (isDelete)
                    {
                        fa._fromRow--;
                        fa._fromCol--;
                    }

                    if (deleteRange.Collide(fa) == ExcelAddressBase.eAddressCollition.Partly)
                    {
                        throw new InvalidOperationException(errorMsg);
                    }
                }

                _ = hs.Add(si);
            }
        }
    }
}