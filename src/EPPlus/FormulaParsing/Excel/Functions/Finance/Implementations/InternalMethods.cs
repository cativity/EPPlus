﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/13/2020         EPPlus Software AB       Implemented function (ported to c# from Microsoft.VisualBasic.Financial.vb (MIT))
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class InternalMethods
    {
        internal static FinanceCalcResult<double> PMT_Internal(double Rate, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            //       Checking for error conditions
            if (NPer == 0.0)
            {
                return new FinanceCalcResult<double>(eErrorType.Value);
            }

            if(Rate == 0.0)
            {
                return new FinanceCalcResult<double>((-FV - PV) / NPer);
            }
            else
            {
                double dTemp;

                if (Due != 0)
                {
                    dTemp = 1.0 + Rate;
                }
                else
                {
                    dTemp = 1.0;
                }

                double dTemp3 = Rate + 1.0;

                //       WARSI Using the exponent operator for pow(..) in C code of PMT. Still got
                //       to make sure that they (pow and ^) are same for all conditions
                double dTemp2 = System.Math.Pow(dTemp3, NPer);
                double result = ((-FV - PV * dTemp2) / (dTemp * (dTemp2 - 1.0)) * Rate);
                return new FinanceCalcResult<double>(result);
            }
        }

        internal static double FV_Internal(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            double dTemp;

            //Performing calculation
            if (Rate == 0)
            {
                return (-PV - Pmt * NPer);
            }

            if (Due != PmtDue.EndOfPeriod)
            {
                dTemp = 1.0 + Rate;
            }
            else
            {
                dTemp = 1.0;
            }

            double dTemp3 = 1.0 + Rate;
            double dTemp2 = System.Math.Pow(dTemp3, NPer);

            //Do divides before multiplies to avoid OverflowExceptions
            return ((-PV) * dTemp2) - ((Pmt / Rate) * dTemp * (dTemp2 - 1.0));
        }
    }
}
