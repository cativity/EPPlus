/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/

using OfficeOpenXml.Core.Worksheet.Fill;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
    #region FillNumbers

    /// <summary>
    /// Fills the range by adding 1 to each cell starting from the value in the top left cell by column
    /// </summary>
    public void FillNumber()
    {
        this.FillNumber(x => { });
    }

    /// <summary>
    /// Fills a range by adding the step value to the start Value. If <paramref name="startValue"/> is null the first value in the row/column is used.
    /// Fill is done by column from top to bottom
    /// </summary>
    /// <param name="startValue">The start value of the first cell. If this value is null the value of the first cell is used.</param>
    /// <param name="stepValue">The value used for each step</param>
    public void FillNumber(double? startValue, double stepValue = 1)
    {
        this.FillNumber(x =>
        {
            x.StepValue = stepValue;
            x.StartValue = startValue;
        });
    }

    /// <summary>
    /// Fills a range by using the argument options. 
    /// </summary>
    /// <param name="options">The option to configure the fill.</param>
    public void FillNumber(Action<FillNumberParams> options)
    {
        FillNumberParams? o = new FillNumberParams();
        options?.Invoke(o);

        if (o.Direction == eFillDirection.Column)
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int c = this._fromCol; c <= this._toCol; c++)
                {
                    FillMethods.FillNumber(this._worksheet, this._fromRow, this._toRow, c, c, o);
                }
            }
            else
            {
                for (int c = this._toCol; c >= this._fromCol; c--)
                {
                    FillMethods.FillNumber(this._worksheet, this._fromRow, this._toRow, c, c, o);
                }
            }
        }
        else
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int r = this._fromRow; r <= this._toRow; r++)
                {
                    FillMethods.FillNumber(this._worksheet, r, r, this._fromCol, this._toCol, o);
                }
            }
            else
            {
                for (int r = this._toRow; r >= this._fromRow; r--)
                {
                    FillMethods.FillNumber(this._worksheet, r, r, this._fromCol, this._toCol, o);
                }
            }
        }

        if (!string.IsNullOrEmpty(o.NumberFormat))
        {
            this.Style.Numberformat.Format = o.NumberFormat;
        }
    }

    #endregion

    #region FillDateTime

    /// <summary>
    /// Fills the range by adding 1 day to each cell starting from the value in the top left cell by column.
    /// </summary>
    public void FillDateTime()
    {
        this.FillDateTime(x => { });
    }

    /// <summary>
    /// Fills the range by adding 1 day to each cell per column starting from <paramref name="startValue"/>.
    /// </summary>
    public void FillDateTime(DateTime? startValue, eDateTimeUnit dateTimeUnit = eDateTimeUnit.Day, int stepValue = 1)
    {
        this.FillDateTime(x =>
        {
            x.StartValue = startValue;
            x.DateTimeUnit = dateTimeUnit;
            x.StepValue = stepValue;
        });
    }

    /// <summary>
    /// Fill the range with dates.
    /// </summary>
    /// <param name="options">Options how to perform the fill</param>
    public void FillDateTime(Action<FillDateParams> options)
    {
        FillDateParams? o = new FillDateParams();
        options?.Invoke(o);

        if (o.Direction == eFillDirection.Column)
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int c = this._fromCol; c <= this._toCol; c++)
                {
                    FillMethods.FillDateTime(this._worksheet, this._fromRow, this._toRow, c, c, o);
                }
            }
            else
            {
                for (int c = this._toCol; c >= this._fromCol; c--)
                {
                    FillMethods.FillDateTime(this._worksheet, this._fromRow, this._toRow, c, c, o);
                }
            }
        }
        else
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int r = this._fromRow; r <= this._toRow; r++)
                {
                    FillMethods.FillDateTime(this._worksheet, r, r, this._fromCol, this._toCol, o);
                }
            }
            else
            {
                for (int r = this._toRow; r >= this._fromRow; r--)
                {
                    FillMethods.FillDateTime(this._worksheet, r, r, this._fromCol, this._toCol, o);
                }
            }
        }

        if (!string.IsNullOrEmpty(o.NumberFormat))
        {
            this.Style.Numberformat.Format = o.NumberFormat;
        }
    }

    #endregion

    #region FillList

    /// <summary>
    /// Fills the range columnwise using the values in the list. 
    /// </summary>
    /// <typeparam name="T">Type used in the list.</typeparam>
    /// <param name="list">The list to use.</param>
    public void FillList<T>(IEnumerable<T> list)
    {
        this.FillList(list, x => { });
    }

    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="list"></param>
    /// <param name="options"></param>
    public void FillList<T>(IEnumerable<T> list, Action<FillListParams> options)
    {
        FillListParams? o = new FillListParams();
        options?.Invoke(o);

        if (o.Direction == eFillDirection.Column)
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int c = this._fromCol; c <= this._toCol; c++)
                {
                    FillMethods.FillList(this._worksheet, this._fromRow, this._toRow, c, c, list, o);
                }
            }
            else
            {
                for (int c = this._toCol; c >= this._fromCol; c--)
                {
                    FillMethods.FillList(this._worksheet, this._fromRow, this._toRow, c, c, list, o);
                }
            }
        }
        else
        {
            if (o.StartPosition == eFillStartPosition.TopLeft)
            {
                for (int r = this._fromRow; r <= this._toRow; r++)
                {
                    FillMethods.FillList(this._worksheet, r, r, this._fromCol, this._toCol, list, o);
                }
            }
            else
            {
                for (int r = this._toRow; r >= this._fromRow; r--)
                {
                    FillMethods.FillList(this._worksheet, r, r, this._fromCol, this._toCol, list, o);
                }
            }
        }

        if (!string.IsNullOrEmpty(o.NumberFormat))
        {
            this.Style.Numberformat.Format = o.NumberFormat;
        }
    }

    #endregion
}