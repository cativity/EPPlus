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
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

/// <summary>
/// Thanks to the guys in this thread: http://stackoverflow.com/questions/2840798/c-sharp-math-class-question
/// </summary>
internal static class MathHelper
{
    // Secant 
    public static double Sec(double x) => 1 / MathObj.Cos(x);

    // Cosecant
    public static double Cosec(double x) => 1 / MathObj.Sin(x);

    // Cotangent 
    public static double Cotan(double x) => 1 / MathObj.Tan(x);

    // Inverse Sine 
    public static double Arcsin(double x) => MathObj.Atan(x / MathObj.Sqrt((-x * x) + 1));

    // Inverse Cosine 
    public static double Arccos(double x) => MathObj.Atan(-x / MathObj.Sqrt((-x * x) + 1)) + (2 * MathObj.Atan(1));

    // Inverse Secant 
    public static double Arcsec(double x) => (2 * MathObj.Atan(1)) - MathObj.Atan(MathObj.Sign(x) / MathObj.Sqrt((x * x) - 1));

    // Inverse Cosecant 
    public static double Arccosec(double x) => MathObj.Atan(MathObj.Sign(x) / MathObj.Sqrt((x * x) - 1));

    // Inverse Cotangent 
    public static double Arccotan(double x) => (2 * MathObj.Atan(1)) - MathObj.Atan(x);

    // Hyperbolic Sine 
    public static double HSin(double x) => (MathObj.Exp(x) - MathObj.Exp(-x)) / 2;

    // Hyperbolic Cosine 
    public static double HCos(double x) => (MathObj.Exp(x) + MathObj.Exp(-x)) / 2;

    // Hyperbolic Tangent 
    public static double HTan(double x) => (MathObj.Exp(x) - MathObj.Exp(-x)) / (MathObj.Exp(x) + MathObj.Exp(-x));

    // Hyperbolic Secant 
    public static double HSec(double x) => 2 / (MathObj.Exp(x) + MathObj.Exp(-x));

    // Hyperbolic Cosecant 
    public static double HCosec(double x) => 2 / (MathObj.Exp(x) - MathObj.Exp(-x));

    // Hyperbolic Cotangent 
    public static double HCotan(double x) => (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));

    // Inverse Hyperbolic Sine 
    public static double HArcsin(double x) => MathObj.Log(x + MathObj.Sqrt((x * x) + 1));

    // Inverse Hyperbolic Cosine 
    public static double HArccos(double x) => MathObj.Log(x + MathObj.Sqrt((x * x) - 1));

    // Inverse Hyperbolic Tangent 
    public static double HArctan(double x) => MathObj.Log((1 + x) / (1 - x)) / 2;

    // Inverse Hyperbolic Secant 
    public static double HArcsec(double x) => MathObj.Log((MathObj.Sqrt((-x * x) + 1) + 1) / x);

    // Inverse Hyperbolic Cosecant 
    public static double HArccosec(double x) => MathObj.Log(((MathObj.Sign(x) * MathObj.Sqrt((x * x) + 1)) + 1) / x);

    // Inverse Hyperbolic Cotangent 
    public static double HArccotan(double x) => MathObj.Log((x + 1) / (x - 1)) / 2;

    // Logarithm to base N 
    public static double LogN(double x, double n) => MathObj.Log(x) / MathObj.Log(n);

    public static double Factorial(double number) => Factorial(number, 1d);

    public static double Factorial(double number, double devisor)
    {
        double result = 1d;

        for (double x = number; x > devisor; x--)
        {
            result *= x;
        }

        return result;
    }

    public static double Radians(double angle) => angle / 180 * MathObj.PI;

    public static int GreatestCommonDevisor(int[] numbers) => numbers.Aggregate(GreatestCommonDevisor);

    static int GreatestCommonDevisor(int a, int b)
    {
        while (a != 0 && b != 0)
        {
            if (a > b)
            {
                a %= b;
            }
            else
            {
                b %= a;
            }
        }

        return a == 0 ? b : a;
    }

    public static int LeastCommonMultiple(int[] numbers) => numbers.Aggregate(LeastCommonMultiple);

    static int LeastCommonMultiple(int a, int b) => a * b / GreatestCommonDevisor(a, b);
}