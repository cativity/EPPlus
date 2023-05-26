/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal
{
    internal static class SortItemLeftToRightFactory
    {
        internal static List<SortItemLeftToRight<ExcelValue>> Create(ExcelRangeBase range)
        {
            List<SortItemLeftToRight<ExcelValue>>? sortItems = new List<SortItemLeftToRight<ExcelValue>>();
            int nRows = range._toRow - range._fromRow + 1;
            int col = range._fromCol;

            while (col <= range._toCol)
            {
                int currentRow = 0;
                SortItemLeftToRight<ExcelValue>? sortItem = new SortItemLeftToRight<ExcelValue> { Column = col, Items = new ExcelValue[nRows] };
                while(currentRow < nRows)
                {
                    int row = currentRow + range._fromRow;
                    ExcelRange? cell = range.Worksheet.Cells[row, col, row, col];
                    ExcelValue v = new ExcelValue();
                    v._styleId = cell.StyleID;
                    v._value = cell.Value;
                    sortItem.Items[currentRow] = v;
                    currentRow++;
                }
                sortItems.Add(sortItem);
                col++;
            }
            return sortItems;
        }
    }
}
