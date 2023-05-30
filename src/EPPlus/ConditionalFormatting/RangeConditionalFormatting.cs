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
using OfficeOpenXml.Utils;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

internal class RangeConditionalFormatting : IRangeConditionalFormatting
{
    #region Public Properties

    public ExcelWorksheet _worksheet;
    public ExcelAddress _address;

    #endregion Public Properties

    #region Constructors

    public RangeConditionalFormatting(ExcelWorksheet worksheet, ExcelAddress address)
    {
        Require.Argument(worksheet).IsNotNull("worksheet");
        Require.Argument(address).IsNotNull("address");

        this._worksheet = worksheet;
        this._address = address;
    }

    #endregion Constructors

    #region Conditional Formatting Rule Types

    /// <summary>
    /// Add AboveOrEqualAverage Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddAboveAverage() => this._worksheet.ConditionalFormatting.AddAboveAverage(this._address);

    /// <summary>
    /// Add AboveOrEqualAverage Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage() => this._worksheet.ConditionalFormatting.AddAboveOrEqualAverage(this._address);

    /// <summary>
    /// Add BelowOrEqualAverage Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddBelowAverage() => this._worksheet.ConditionalFormatting.AddBelowAverage(this._address);

    /// <summary>
    /// Add BelowOrEqualAverage Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage() => this._worksheet.ConditionalFormatting.AddBelowOrEqualAverage(this._address);

    /// <summary>
    /// Add AboveStdDev Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingStdDevGroup AddAboveStdDev() => this._worksheet.ConditionalFormatting.AddAboveStdDev(this._address);

    /// <summary>
    /// Add BelowStdDev Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingStdDevGroup AddBelowStdDev() => this._worksheet.ConditionalFormatting.AddBelowStdDev(this._address);

    /// <summary>
    /// Add Bottom Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddBottom() => this._worksheet.ConditionalFormatting.AddBottom(this._address);

    /// <summary>
    /// Add BottomPercent Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddBottomPercent() => this._worksheet.ConditionalFormatting.AddBottomPercent(this._address);

    /// <summary>
    /// Add Top Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddTop() => this._worksheet.ConditionalFormatting.AddTop(this._address);

    /// <summary>
    /// Add TopPercent Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddTopPercent() => this._worksheet.ConditionalFormatting.AddTopPercent(this._address);

    /// <summary>
    /// Add Last7Days Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLast7Days() => this._worksheet.ConditionalFormatting.AddLast7Days(this._address);

    /// <summary>
    /// Add LastMonth Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLastMonth() => this._worksheet.ConditionalFormatting.AddLastMonth(this._address);

    /// <summary>
    /// Add LastWeek Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLastWeek() => this._worksheet.ConditionalFormatting.AddLastWeek(this._address);

    /// <summary>
    /// Add NextMonth Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddNextMonth() => this._worksheet.ConditionalFormatting.AddNextMonth(this._address);

    /// <summary>
    /// Add NextWeek Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddNextWeek() => this._worksheet.ConditionalFormatting.AddNextWeek(this._address);

    /// <summary>
    /// Add ThisMonth Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddThisMonth() => this._worksheet.ConditionalFormatting.AddThisMonth(this._address);

    /// <summary>
    /// Add ThisWeek Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddThisWeek() => this._worksheet.ConditionalFormatting.AddThisWeek(this._address);

    /// <summary>
    /// Add Today Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddToday() => this._worksheet.ConditionalFormatting.AddToday(this._address);

    /// <summary>
    /// Add Tomorrow Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddTomorrow() => this._worksheet.ConditionalFormatting.AddTomorrow(this._address);

    /// <summary>
    /// Add Yesterday Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddYesterday() => this._worksheet.ConditionalFormatting.AddYesterday(this._address);

    /// <summary>
    /// Add BeginsWith Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingBeginsWith AddBeginsWith() => this._worksheet.ConditionalFormatting.AddBeginsWith(this._address);

    /// <summary>
    /// Add Between Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingBetween AddBetween() => this._worksheet.ConditionalFormatting.AddBetween(this._address);

    /// <summary>
    /// Add ContainsBlanks Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingContainsBlanks AddContainsBlanks() => this._worksheet.ConditionalFormatting.AddContainsBlanks(this._address);

    /// <summary>
    /// Add ContainsErrors Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingContainsErrors AddContainsErrors() => this._worksheet.ConditionalFormatting.AddContainsErrors(this._address);

    /// <summary>
    /// Add ContainsText Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingContainsText AddContainsText() => this._worksheet.ConditionalFormatting.AddContainsText(this._address);

    /// <summary>
    /// Add DuplicateValues Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingDuplicateValues AddDuplicateValues() => this._worksheet.ConditionalFormatting.AddDuplicateValues(this._address);

    /// <summary>
    /// Add EndsWith Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingEndsWith AddEndsWith() => this._worksheet.ConditionalFormatting.AddEndsWith(this._address);

    /// <summary>
    /// Add Equal Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingEqual AddEqual() => this._worksheet.ConditionalFormatting.AddEqual(this._address);

    /// <summary>
    /// Add Expression Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingExpression AddExpression() => this._worksheet.ConditionalFormatting.AddExpression(this._address);

    /// <summary>
    /// Add GreaterThan Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingGreaterThan AddGreaterThan() => this._worksheet.ConditionalFormatting.AddGreaterThan(this._address);

    /// <summary>
    /// Add GreaterThanOrEqual Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual() => this._worksheet.ConditionalFormatting.AddGreaterThanOrEqual(this._address);

    /// <summary>
    /// Add LessThan Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingLessThan AddLessThan() => this._worksheet.ConditionalFormatting.AddLessThan(this._address);

    /// <summary>
    /// Add LessThanOrEqual Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual() => this._worksheet.ConditionalFormatting.AddLessThanOrEqual(this._address);

    /// <summary>
    /// Add NotBetween Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingNotBetween AddNotBetween() => this._worksheet.ConditionalFormatting.AddNotBetween(this._address);

    /// <summary>
    /// Add NotContainsBlanks Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks() => this._worksheet.ConditionalFormatting.AddNotContainsBlanks(this._address);

    /// <summary>
    /// Add NotContainsErrors Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors() => this._worksheet.ConditionalFormatting.AddNotContainsErrors(this._address);

    /// <summary>
    /// Add NotContainsText Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingNotContainsText AddNotContainsText() => this._worksheet.ConditionalFormatting.AddNotContainsText(this._address);

    /// <summary>
    /// Add NotEqual Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingNotEqual AddNotEqual() => this._worksheet.ConditionalFormatting.AddNotEqual(this._address);

    /// <summary>
    /// Add UniqueValues Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingUniqueValues AddUniqueValues() => this._worksheet.ConditionalFormatting.AddUniqueValues(this._address);

    /// <summary>
    /// Add ThreeColorScale Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingThreeColorScale AddThreeColorScale() =>
        (IExcelConditionalFormattingThreeColorScale)this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.ThreeColorScale,
                                                                                                  this._address);

    /// <summary>
    /// Add TwoColorScale Conditional Formatting
    /// </summary>
    ///  <returns></returns>
    public IExcelConditionalFormattingTwoColorScale AddTwoColorScale() =>
        (IExcelConditionalFormattingTwoColorScale)this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.TwoColorScale,
                                                                                                this._address);

    /// <summary>
    /// Adds a ThreeIconSet rule 
    /// </summary>
    /// <param name="IconSet"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(eExcelconditionalFormatting3IconsSetType IconSet)
    {
        IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>? rule =
            (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)
            this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.ThreeIconSet, this._address);

        rule.IconSet = IconSet;

        return rule;
    }

    /// <summary>
    /// Adds a FourIconSet rule 
    /// </summary>
    /// <param name="IconSet"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(eExcelconditionalFormatting4IconsSetType IconSet)
    {
        IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>? rule =
            (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)
            this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.FourIconSet, this._address);

        rule.IconSet = IconSet;

        return rule;
    }

    /// <summary>
    /// Adds a FiveIconSet rule 
    /// </summary>
    /// <param name="IconSet"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(eExcelconditionalFormatting5IconsSetType IconSet)
    {
        IExcelConditionalFormattingFiveIconSet? rule =
            (IExcelConditionalFormattingFiveIconSet)this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.FiveIconSet,
                                                                                                  this._address);

        rule.IconSet = IconSet;

        return rule;
    }

    /// <summary>
    /// Adds a Databar rule 
    /// </summary>
    /// <param name="Color">The color of the databar</param>
    /// <returns></returns>
    public IExcelConditionalFormattingDataBarGroup AddDatabar(System.Drawing.Color Color)
    {
        IExcelConditionalFormattingDataBarGroup? rule =
            (IExcelConditionalFormattingDataBarGroup)this._worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.DataBar, this._address);

        rule.Color = Color;

        return rule;
    }

    #endregion Conditional Formatting Rule Types
}