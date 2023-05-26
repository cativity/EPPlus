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

namespace OfficeOpenXml.FormulaParsing
{
    public class ExcelCell
    {
        public ExcelCell(object val, string formula, int colIndex, int rowIndex)
        {
            this.Value = val;
            this.Formula = formula;
            this.ColIndex = colIndex;
            this.RowIndex = rowIndex;
        }

        public int ColIndex { get; private set; }

        public int RowIndex { get; private set; }

        public object Value { get; private set; }

        public string Formula { get; private set; }
    }
}
