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

using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing;

internal class DependencyChain
{
    internal List<FormulaCell> list = new List<FormulaCell>();
    internal Dictionary<ulong, int> index = new Dictionary<ulong, int>();
    internal List<int> CalcOrder = new List<int>();

    internal void Add(FormulaCell f)
    {
        this.list.Add(f);
        f.Index = this.list.Count - 1;
        this.index.Add(ExcelCellBase.GetCellId(f.wsIndex, f.Row, f.Column), f.Index);
    }
}