/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function (ported to c# from Microsoft.VisualBasic.Financial.vb (MIT))
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class MirrImpl
    {
        internal static double LDoNPV(double Rate, ref double[] ValueArray, int iWNType)
        {
            bool bSkipPos = iWNType < 0;
            bool bSkipNeg = iWNType > 0;

            double dTemp2 = 1.0;
            double dTotal = 0.0;

            int lLower = 0;
            int lUpper = ValueArray.Length - 1;

            for(int I = lLower; I <= lUpper; I++)
            {
                double dTVal = ValueArray[I];
                dTemp2 += dTemp2 * Rate;

                if(!((bSkipPos && dTVal > 0.0) || (bSkipNeg && dTVal< 0.0)))
                {
                    dTotal += dTVal / dTemp2;
                }
           
            }
            return dTotal;
        }

        internal static FinanceCalcResult<double> MIRR(double[] ValueArray, double FinanceRate, double ReinvestRate)
        {
            if(ValueArray.Rank != 1)
            {
                return new FinanceCalcResult<double>(eErrorType.Value);
            }

            int lLower = 0;
            int lUpper = ValueArray.Length - 1;
            int lCVal = lUpper - lLower + 1;

            if(FinanceRate == -1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if(ReinvestRate == -1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if(lCVal <= 1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            double dNpvNeg = LDoNPV(FinanceRate, ref ValueArray, -1);

            if (dNpvNeg == 0.0)
            {
                return new FinanceCalcResult<double>(eErrorType.Div0);
            }

            double dNpvPos = LDoNPV(ReinvestRate, ref ValueArray, 1); // npv of +ve values
            double dTemp1 = ReinvestRate + 1.0;
            double dNTemp2 = lCVal;

            double dTemp = -dNpvPos * System.Math.Pow(dTemp1, dNTemp2) / (dNpvNeg * (FinanceRate + 1.0));

            if (dTemp < 0d)
            {
                return new FinanceCalcResult<double>(eErrorType.Value);
            }

            dTemp1 = 1d / (lCVal - 1d);

            return new FinanceCalcResult<double>(System.Math.Pow(dTemp, dTemp1) - 1.0);
        }

    }
}
