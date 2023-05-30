using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;

internal class YieldImpl
{
    public YieldImpl(ICouponProvider couponProvider, IPriceProvider priceProvider)
    {
        this._couponProvider = couponProvider;
        this._priceProvider = priceProvider;
    }

    private readonly ICouponProvider _couponProvider;
    private readonly IPriceProvider _priceProvider;

    private static bool AreEqual(double x, double y) => System.Math.Abs(x - y) < 0.000000001;

    public double GetYield(System.DateTime settlement,
                           System.DateTime maturity,
                           double rate,
                           double pr,
                           double redemption,
                           int frequency,
                           DayCountBasis basis = DayCountBasis.US_30_360)
    {
        double A = this._couponProvider.GetCoupdaybs(settlement, maturity, frequency, basis);
        double N = this._couponProvider.GetCoupnum(settlement, maturity, frequency, basis);
        double E = this._couponProvider.GetCoupdays(settlement, maturity, frequency, basis);

        if (N <= -1)
        {
            double DSR = E - A;
            double part1 = (redemption / 100) + (rate / frequency) - ((pr / 100d) + (A / E * rate / frequency));
            double part2 = (pr / 100d) + (A / E * rate / frequency);
            double retVal = part1 / part2 * (frequency * E / DSR);

            return retVal;
        }
        else
        {
            double fPriceN = 0.0;
            double fYield1 = 0.0;
            double fYield2 = 1.0;
            double fPrice1 = this._priceProvider.GetPrice(settlement, maturity, rate, fYield1, redemption, frequency, basis);
            double fPrice2 = this._priceProvider.GetPrice(settlement, maturity, rate, fYield2, redemption, frequency, basis);
            double fYieldN = (fYield2 - fYield1) * 0.5;

            for (int nIter = 0; nIter < 100 && !AreEqual(fPriceN, pr); nIter++)
            {
                fPriceN = this._priceProvider.GetPrice(settlement, maturity, rate, fYieldN, redemption, frequency, basis);

                if (AreEqual(pr, fPrice1))
                {
                    return fYield1;
                }
                else if (AreEqual(pr, fPrice2))
                {
                    return fYield2;
                }
                else if (AreEqual(pr, fPriceN))
                {
                    return fYieldN;
                }
                else if (pr < fPrice2)
                {
                    fYield2 *= 2.0;
                    fPrice2 = this._priceProvider.GetPrice(settlement, maturity, rate, fYield2, redemption, frequency, basis);

                    fYieldN = (fYield2 - fYield1) * 0.5;
                }
                else
                {
                    if (pr < fPriceN)
                    {
                        fYield1 = fYieldN;
                        fPrice1 = fPriceN;
                    }
                    else
                    {
                        fYield2 = fYieldN;
                        fPrice2 = fPriceN;
                    }

                    fYieldN = fYield2 - ((fYield2 - fYield1) * ((pr - fPrice2) / (fPrice1 - fPrice2)));
                }
            }

            if (System.Math.Abs(pr - fPriceN) > pr / 100d)
            {
                throw new Exception("Result not precise enough");
            }

            return fYieldN;
        }
    }
}