/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal abstract class RankFunctionBase : ExcelFunction
{
    protected static List<double> GetNumbersFromRange(FunctionArgument refArg, bool sortAscending)
    {
        List<double>? numbers = new List<double>();
        foreach (ICellInfo? cell in refArg.ValueAsRangeInfo)
        {
            double cellValue = Utils.ConvertUtil.GetValueDouble(cell.Value, false, true);
            if (!double.IsNaN(cellValue))
            {
                numbers.Add(cellValue);
            }
        }
        if (sortAscending)
        {
            numbers.Sort();
        }
        else
        {
            numbers.Sort((x, y) => y.CompareTo(x));
        }

        return numbers;
    }

    protected double[] GetNumbersFromArgs(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
    {
        double[]? array = this.ArgsToDoubleEnumerable(new FunctionArgument[] { arguments.ElementAt(index) }, context)
                              .Select(x => (double)x)
                              .OrderBy(x => x)
                              .ToArray();
        return array;
    }

    protected static double PercentRankIncImpl(double[] array, double number)
    {
        double smallerThan = 0d;
        double largestBelow = 0d;
        int ix = 0;
        while (number > array[ix])
        {
            smallerThan++;
            largestBelow = array[ix];
            ix++;
        }
        bool fullMatch = AreEqual(number, array[ix]);
        while (ix < array.Length - 1 && AreEqual(number, array[ix]))
        {
            ix++;
        }

        double smallestAbove = array[ix];
        int largerThan = AreEqual(number, array[array.Length - 1]) ? 0 : array.Length - ix;
        if (fullMatch)
        {
            return smallerThan / (smallerThan + largerThan);
        }

        double percentrankLow = PercentRankIncImpl(array, largestBelow);
        double percentrankHigh = PercentRankIncImpl(array, smallestAbove);
        return percentrankLow + (percentrankHigh - percentrankLow) * ((number - largestBelow) / (smallestAbove - largestBelow));
    }

    protected static double PercentRankExcImpl(double[] array, double number)
    {
        double smallerThan = 0d;
        double largestBelow = 0d;
        int ix = 0;
        while (number > array[ix])
        {
            smallerThan++;
            largestBelow = array[ix];
            ix++;
        }
        smallerThan++;
        bool fullMatch = AreEqual(number, array[ix]);
        while (ix < array.Length - 1 && AreEqual(number, array[ix]))
        {
            ix++;
        }

        double smallestAbove = array[ix];
        int largerThan = AreEqual(number, array[array.Length - 1]) ? 0 : array.Length - ix + 1;
        if (fullMatch)
        {
            return smallerThan / (smallerThan + largerThan);
        }

        double percentrankLow = PercentRankExcImpl(array, largestBelow);
        double percentrankHigh = PercentRankExcImpl(array, smallestAbove);
        return percentrankLow + (percentrankHigh - percentrankLow) * ((number - largestBelow) / (smallestAbove - largestBelow));
    }

    /// <summary>
    /// Rank functions rounds towards zero, i.e. 0.41666666 should be rounded to 0.4166 if 4 decimals.
    /// </summary>
    /// <param name="number">The number to round</param>
    /// <param name="sign">Number of siginicant digits</param>
    /// <returns></returns>
    protected static double RoundResult(double number, int sign)
    {
        return RoundingHelper.RoundToSignificantFig(number, sign, false);
    }
}