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
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal static class RoundingHelper
{
    public enum Direction
    {
        Up,
        Down,
        AlwaysDown,
        AlwaysUp,
        Nearest
    }

    public static double Round(double number, double multiple, Direction direction)
    {
        if (multiple == 0)
        {
            return 0d;
        }

        bool isNegativeNumber = number < 0;
        bool isNegativeMultiple = multiple < 0;
        double n = isNegativeNumber ? number * -1 : number;
        double m = isNegativeMultiple ? multiple * -1 : multiple;
        if (number % multiple == 0)
        {
            return number;
        }
        else if (multiple > -1 && multiple < 1)
        {

            double floor = System.Math.Floor(n);
            double rest = n - floor;
            int nSign = (int)(rest / m) + 1;

            double upperRound = System.Math.Round(nSign * m, 14);
            double lowerRound = System.Math.Round((nSign - 1) * m, 14);
            return ExecuteRounding(n, floor + lowerRound, floor + upperRound, direction, isNegativeNumber);
        }
        double result = double.NaN;
        if (m == 1)
        {
            if (direction == Direction.Up || direction == Direction.AlwaysUp)
            {
                if (direction == Direction.AlwaysUp && isNegativeNumber)
                {
                    result = System.Math.Floor(n);
                }
                else
                {
                    result = System.Math.Ceiling(n);
                }
            }
            else if (direction == Direction.Nearest)
            {
                result = System.Math.Floor(n);
                if (n % 1 >= 0.5)
                {
                    result++;
                }
            }
            else
            {
                if (direction == Direction.AlwaysDown && isNegativeNumber)
                {
                    result = System.Math.Ceiling(n);
                }
                else
                {
                    result = System.Math.Floor(n);
                }
            }
        }
        else if (m > n)
        {
            return ExecuteRounding(n, 0, m, direction, isNegativeNumber);
        }
        else if (direction == Direction.Up || direction == Direction.AlwaysUp)
        {
            if (direction == Direction.AlwaysUp && number < 0)
            {
                if (multiple < 0)
                {
                    multiple *= -1;
                }

                return System.Math.Round(number - (number % multiple), 14);
            }
            return System.Math.Round(number - (number % multiple) + multiple, 14);
        }
        else if (direction == Direction.Nearest)
        {
            if (n % m >= m / 2d)
            {
                result = System.Math.Round(n + (m - (n % m)));
            }
            else
            {
                result = System.Math.Round(n - (n % m));
            }
        }
        else
        {
            if (direction == Direction.AlwaysDown && number < 0)
            {
                if (multiple < 0)
                {
                    multiple *= -1;
                }

                return System.Math.Round(number - (number % multiple) - multiple, 14);
            }

            return System.Math.Round(number - (number % multiple), 14);
        }
        return isNegativeNumber ? -1 * result : result;
    }

    public static double ExecuteRounding(double number, double lowerRound, double upperRound, Direction direction, bool isNegativeNumber)
    {
        double result = double.NaN;
        if (direction == Direction.Nearest)
        {
            if (upperRound - number > number - lowerRound)
            {
                result = lowerRound;
            }
            else
            {
                result = upperRound;
            }
        }
        else if (direction == Direction.AlwaysUp)
        {
            result = isNegativeNumber ? lowerRound : upperRound;
        }
        else if (direction == Direction.Up)
        {
            result = upperRound;
        }
        else if (direction == Direction.AlwaysDown)
        {
            result = isNegativeNumber ? upperRound : lowerRound;
        }
        else
        {
            result = lowerRound;
        }
        return isNegativeNumber ? -1 * result : result;
    }


    internal static bool IsInvalidNumberAndSign(double number, double sign)
    {
        return number > 0d && sign < 0;
    }

    internal static double RoundToSignificantFig(double number, double nSignificantFigures)
    {
        return RoundToSignificantFig(number, nSignificantFigures, true);
    }

    internal static double RoundToSignificantFig(double number, double nSignificantFigures, bool awayFromMidpoint)
    {
        bool isNegative = false;
        if(number < 0d)
        {
            number *= -1;
            isNegative = true;
        }
        double nFiguresIntPart = GetNumberOfDigitsIntPart(number);
        double nLeadingZeroDecimals = GetNumberOfLeadingZeroDecimals(number);
        double nFiguresDecimalPart = nSignificantFigures - nFiguresIntPart - nLeadingZeroDecimals;
        if (number < 1d)
        {
            nFiguresDecimalPart -= nLeadingZeroDecimals;
        }
        double tmp = number * System.Math.Pow(10, nFiguresDecimalPart + nLeadingZeroDecimals);
        double e = awayFromMidpoint? tmp + 0.5 : tmp;
        if(awayFromMidpoint)
        { 
            if ((float)e == (float)System.Math.Ceiling(tmp))
            {
                double f = System.Math.Ceiling(tmp);
                int h = (int)f - 2;
                if (h % 2 != 0)
                {
                    e -= 1;
                }
            }
        }
        double intVersion = System.Math.Floor(e);
        double divideBy = System.Math.Pow(10, nFiguresDecimalPart + nLeadingZeroDecimals);
        double result = intVersion / divideBy;
        return isNegative ? result * -1 : result;
    }

    /// <summary>
    /// Count the number of digits left of the decimal point
    /// </summary>
    /// <param name="n"></param>
    /// <returns></returns>
    internal static double GetNumberOfDigitsIntPart(double n)
    {
        double tmp = n;
        int nFiguresIntPart;
        for (nFiguresIntPart = 0; tmp >= 1; ++nFiguresIntPart)
        {
            tmp /= 10;
        }

        return nFiguresIntPart;
    }

    private static double GetNumberOfLeadingZeroDecimals(double n)
    {
        if (n == 0)
        {
            return 0;
        }

        double tmp = n;
        int result = 0;
        while (tmp < 1d)
        {
            tmp *= 10;
            result++;
        }
        return result - 1;
    }
}