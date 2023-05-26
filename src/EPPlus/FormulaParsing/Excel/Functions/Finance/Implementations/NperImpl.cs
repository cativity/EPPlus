﻿/*************************************************************************************************
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
    public static class NperImpl
    {
        public static FinanceCalcResult<double> NPer(double Rate, double Pmt, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            //   Checking Error Conditions
            if (Rate <= -1.0)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if (Rate == 0.0)
            {
                if (Pmt == 0.0)
                {
                    return new FinanceCalcResult<double>(eErrorType.Num);
                }

                return new FinanceCalcResult<double>(-(PV + FV) / Pmt);
            }
            else
            {
                double dTemp3;

                if (Due != 0)
                {
                    dTemp3 = Pmt * (1.0 + Rate) / Rate;
                }
                else
                {
                    dTemp3 = Pmt / Rate;
                }
                double dTempFv = -FV + dTemp3;
                double dTempPv = PV + dTemp3;

                //       Make sure the values fit the domain of log()
                if( dTempFv< 0.0 && dTempPv < 0.0)
                {
                    dTempFv = -1 * dTempFv;
                    dTempPv = -1 * dTempPv;
                }
                else if(dTempFv <= 0.0 || dTempPv <= 0.0)
                {
                    return new FinanceCalcResult<double>(eErrorType.Num);
                }
                double dTemp4 = Rate + 1.0;

                double result = (System.Math.Log(dTempFv) - System.Math.Log(dTempPv)) / System.Math.Log(dTemp4);
                return new FinanceCalcResult<double>(result);
            }
        }
    }
}
