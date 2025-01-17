﻿using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.ConditionalFormatting;

/// <summary>
/// Provides a simple way to type cast a conditional formatting object to its top level class.
/// </summary>
public class ExcelConditionalFormattingAsType
{
    IExcelConditionalFormattingRule _rule;

    internal ExcelConditionalFormattingAsType(IExcelConditionalFormattingRule rule) => this._rule = rule;

    /// <summary>
    /// Converts the conditional formatting object to it's top level or another nested class.        
    /// </summary>
    /// <typeparam name="T">The type of conditional formatting object. T must be inherited from IExcelConditionalFormattingRule</typeparam>
    /// <returns>The conditional formatting rule as type T</returns>
    public T Type<T>()
        where T : IExcelConditionalFormattingRule
    {
        if (this._rule is T t)
        {
            return t;
        }

        return default;
    }

    /// <summary>
    /// Returns the conditional formatting object as an Average rule
    /// If this object is not of type AboveAverage, AboveOrEqualAverage, BelowAverage or BelowOrEqualAverage, null will be returned
    /// </summary>
    /// <returns>The conditional formatting rule as an Average rule</returns>
    public IExcelConditionalFormattingAverageGroup Average => this._rule as IExcelConditionalFormattingAverageGroup;

    /// <summary>
    /// Returns the conditional formatting object as a StdDev rule
    /// If this object is not of type AboveStdDev or BelowStdDev, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a StdDev rule</returns>
    public IExcelConditionalFormattingStdDevGroup StdDev => this._rule as IExcelConditionalFormattingStdDevGroup;

    /// <summary>
    /// Returns the conditional formatting object as a TopBottom rule
    /// If this object is not of type Bottom, BottomPercent, Top or TopPercent, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a TopBottom rule</returns>
    public IExcelConditionalFormattingTopBottomGroup TopBottom => this._rule as IExcelConditionalFormattingTopBottomGroup;

    /// <summary>
    /// Returns the conditional formatting object as a DateTimePeriod rule
    /// If this object is not of type Last7Days, LastMonth, LastWeek, NextMonth, NextWeek, ThisMonth, ThisWeek, Today, Tomorrow or Yesterday, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a DateTimePeriod rule</returns>
    public IExcelConditionalFormattingTimePeriodGroup DateTimePeriod => this._rule as IExcelConditionalFormattingTimePeriodGroup;

    /// <summary>
    /// Returns the conditional formatting object as a Between rule
    /// If this object is not of type Between, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a Between rule</returns>
    public IExcelConditionalFormattingBetween Between => this._rule as IExcelConditionalFormattingBetween;

    /// <summary>
    /// Returns the conditional formatting object as a ContainsBlanks rule
    /// If this object is not of type ContainsBlanks, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a ContainsBlanks rule</returns>
    public IExcelConditionalFormattingContainsBlanks ContainsBlanks => this._rule as IExcelConditionalFormattingContainsBlanks;

    /// <summary>
    /// Returns the conditional formatting object as a ContainsErrors rule
    /// If this object is not of type ContainsErrors, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a ContainsErrors rule</returns>
    public IExcelConditionalFormattingContainsErrors ContainsErrors => this._rule as IExcelConditionalFormattingContainsErrors;

    /// <summary>
    /// Returns the conditional formatting object as a ContainsText rule
    /// If this object is not of type ContainsText, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a ContainsText rule</returns>
    public IExcelConditionalFormattingContainsText ContainsText => this._rule as IExcelConditionalFormattingContainsText;

    /// <summary>
    /// Returns the conditional formatting object as a NotContainsBlanks rule
    /// If this object is not of type NotContainsBlanks, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a NotContainsBlanks rule</returns>
    public IExcelConditionalFormattingNotContainsBlanks NotContainsBlanks => this._rule as IExcelConditionalFormattingNotContainsBlanks;

    /// <summary>
    /// Returns the conditional formatting object as a NotContainsText rule
    /// If this object is not of type NotContainsText, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a NotContainsText rule</returns>
    public IExcelConditionalFormattingNotContainsText NotContainsText => this._rule as IExcelConditionalFormattingNotContainsText;

    /// <summary>
    /// Returns the conditional formatting object as a NotContainsErrors rule
    /// If this object is not of type NotContainsErrors, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a NotContainsErrors rule</returns>
    public IExcelConditionalFormattingNotContainsErrors NotContainsErrors => this._rule as IExcelConditionalFormattingNotContainsErrors;

    /// <summary>
    /// Returns the conditional formatting object as a NotBetween rule
    /// If this object is not of type NotBetween, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a NotBetween rule</returns>
    public IExcelConditionalFormattingNotBetween NotBetween => this._rule as IExcelConditionalFormattingNotBetween;

    /// <summary>
    /// Returns the conditional formatting object as an Equal rule
    /// If this object is not of type Equal, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as an Equal rule</returns>
    public IExcelConditionalFormattingEqual Equal => this._rule as IExcelConditionalFormattingEqual;

    /// <summary>
    /// Returns the conditional formatting object as a NotEqual rule
    /// If this object is not of type NotEqual, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a NotEqual rule</returns>
    public IExcelConditionalFormattingNotEqual NotEqual => this._rule as IExcelConditionalFormattingNotEqual;

    /// <summary>
    /// Returns the conditional formatting object as a DuplicateValues rule
    /// If this object is not of type DuplicateValues, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a DuplicateValues rule</returns>
    public IExcelConditionalFormattingDuplicateValues DuplicateValues => this._rule as IExcelConditionalFormattingDuplicateValues;

    /// <summary>
    /// Returns the conditional formatting object as a BeginsWith rule
    /// If this object is not of type BeginsWith, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a BeginsWith rule</returns>
    public IExcelConditionalFormattingBeginsWith BeginsWith => this._rule as IExcelConditionalFormattingBeginsWith;

    /// <summary>
    /// Returns the conditional formatting object as an EndsWith rule
    /// If this object is not of type EndsWith, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as an EndsWith rule</returns>
    public IExcelConditionalFormattingEndsWith EndsWith => this._rule as IExcelConditionalFormattingEndsWith;

    /// <summary>
    /// Returns the conditional formatting object as an Expression rule
    /// If this object is not of type Expression, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as an Expression rule</returns>
    public IExcelConditionalFormattingExpression Expression => this._rule as IExcelConditionalFormattingExpression;

    /// <summary>
    /// Returns the conditional formatting object as a GreaterThan rule
    /// If this object is not of type GreaterThan, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a GreaterThan rule</returns>
    public IExcelConditionalFormattingGreaterThan GreaterThan => this._rule as IExcelConditionalFormattingGreaterThan;

    /// <summary>
    /// Returns the conditional formatting object as a GreaterThanOrEqual rule
    /// If this object is not of type GreaterThanOrEqual, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a GreaterThanOrEqual rule</returns>
    public IExcelConditionalFormattingGreaterThanOrEqual GreaterThanOrEqual => this._rule as IExcelConditionalFormattingGreaterThanOrEqual;

    /// <summary>
    /// Returns the conditional formatting object as a LessThan rule
    /// If this object is not of type LessThan, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a LessThan rule</returns>
    public IExcelConditionalFormattingLessThan LessThan => this._rule as IExcelConditionalFormattingLessThan;

    /// <summary>
    /// Returns the conditional formatting object as a LessThanOrEqual rule
    /// If this object is not of type LessThanOrEqual, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a LessThanOrEqual rule</returns>
    public IExcelConditionalFormattingLessThanOrEqual LessThanOrEqual => this._rule as IExcelConditionalFormattingLessThanOrEqual;

    /// <summary>
    /// Returns the conditional formatting object as a UniqueValues rule
    /// If this object is not of type UniqueValues, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a UniqueValues rule</returns>
    public IExcelConditionalFormattingUniqueValues UniqueValues => this._rule as IExcelConditionalFormattingUniqueValues;

    /// <summary>
    /// Returns the conditional formatting object as a TwoColorScale rule
    /// If this object is not of type TwoColorScale, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a TwoColorScale rule</returns>
    public IExcelConditionalFormattingTwoColorScale TwoColorScale => this._rule as IExcelConditionalFormattingTwoColorScale;

    /// <summary>
    /// Returns the conditional formatting object as a ThreeColorScale rule
    /// If this object is not of type ThreeColorScale, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a ThreeColorScale rule</returns>
    public IExcelConditionalFormattingThreeColorScale ThreeColorScale => this._rule as IExcelConditionalFormattingThreeColorScale;

    /// <summary>
    /// Returns the conditional formatting object as a ThreeIconSet rule
    /// If this object is not of type ThreeIconSet, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a ThreeIconSet rule</returns>
    public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> ThreeIconSet => this._rule as IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>;

    /// <summary>
    /// Returns the conditional formatting object as a FourIconSet rule
    /// If this object is not of type FourIconSet, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a FourIconSet rule</returns>
    public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> FourIconSet => this._rule as IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>;

    /// <summary>
    /// Returns the conditional formatting object as a FiveIconSet rule
    /// If this object is not of type FiveIconSet, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a FiveIconSet rule</returns>
    public IExcelConditionalFormattingFiveIconSet FiveIconSet => this._rule as IExcelConditionalFormattingFiveIconSet;

    /// <summary>
    /// Returns the conditional formatting object as a DataBar rule
    /// If this object is not of type DataBar, null will be returned
    /// </summary>
    /// <returns>The conditional formatting object as a DataBar rule</returns>
    public IExcelConditionalFormattingDataBarGroup DataBar => this._rule as IExcelConditionalFormattingDataBarGroup;
}