using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class RangeCopyStylesHelper
    {
        private readonly ExcelRangeBase _sourceRange;
        private readonly ExcelRangeBase _destinationRange;
        internal RangeCopyStylesHelper(ExcelRangeBase sourceRange, ExcelRangeBase destinationRange)
        {
            this._sourceRange = sourceRange;
            this._destinationRange = destinationRange;
        }
        internal void CopyStyles()
        {
            Dictionary<int, int>? styleCashe = new Dictionary<int, int>();
            ExcelWorksheet? wsSource = this._sourceRange.Worksheet;
            ExcelWorksheet? wsDest= this._destinationRange.Worksheet;
            bool sameWorkbook = wsSource.Workbook == wsDest.Workbook; 
            int sc = this._sourceRange._fromCol;
            for(int dc= this._destinationRange._fromCol; dc <= this._destinationRange._toCol; dc++)
            {
                int sr = this._sourceRange._fromRow;
                for (int dr = this._destinationRange._fromRow; dr <= this._destinationRange._toRow; dr++)
                {
                    int styleId = GetStyleId(wsSource, sc, sr);
                    if (!sameWorkbook)
                    {
                        if (styleCashe.ContainsKey(styleId))
                        {
                            styleId = styleCashe[styleId];
                        }
                        else
                        {
                            int sourceStyleId = styleId;
                            styleId = wsDest.Workbook.Styles.CloneStyle(wsSource.Workbook.Styles, styleId);
                            styleCashe.Add(sourceStyleId, styleId);
                        }
                    }

                    this._destinationRange.Worksheet.SetStyleInner(dr, dc, styleId);

                    if (sr < this._sourceRange._toRow)
                    {
                        sr++;
                    }
                }
                if (sc < this._sourceRange._toCol)
                {
                    sc++;
                }
            }
        }

        private static int GetStyleId(ExcelWorksheet wsSource, int sc, int sr)
        {
            int styleId = wsSource.GetStyleInner(sr, sc);
            if (styleId == 0)
            {
                styleId = wsSource.GetStyleInner(sr, 0);
                if (styleId == 0)
                {
                    styleId = wsSource.GetStyleInner(0, sc);
                }
            }

            return styleId;
        }
    }
}
