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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class BuiltInFunctions : FunctionsModule
{
    /// <summary>
    /// 
    /// </summary>
    public BuiltInFunctions()
    {
        // Text
        this.Functions["len"] = new Len();
        this.Functions["lower"] = new Lower();
        this.Functions["upper"] = new Upper();
        this.Functions["left"] = new Left();
        this.Functions["right"] = new Right();
        this.Functions["mid"] = new Mid();
        this.Functions["replace"] = new Replace();
        this.Functions["rept"] = new Rept();
        this.Functions["substitute"] = new Substitute();
        this.Functions["concatenate"] = new Concatenate();
        this.Functions["concat"] = new Concat();
        this.Functions["textjoin"] = new Textjoin();
        this.Functions["char"] = new CharFunction();
        this.Functions["exact"] = new Exact();
        this.Functions["find"] = new Find();
        this.Functions["fixed"] = new Fixed();
        this.Functions["proper"] = new Proper();
        this.Functions["search"] = new Search();
        this.Functions["text"] = new Text.Text();
        this.Functions["t"] = new T();
        this.Functions["hyperlink"] = new Hyperlink();
        this.Functions["value"] = new Value(CultureInfo.CurrentCulture);
        this.Functions["trim"] = new Trim();
        this.Functions["clean"] = new Clean();
        this.Functions["unicode"] = new Unicode();
        this.Functions["unichar"] = new Unichar();
        this.Functions["numbervalue"] = new NumberValue();
        this.Functions["dollar"] = new Dollar();

        // Numbers
        this.Functions["int"] = new CInt();

        // Math
        this.Functions["aggregate"] = new Aggregate();
        this.Functions["abs"] = new Abs();
        this.Functions["asin"] = new Asin();
        this.Functions["asinh"] = new Asinh();
        this.Functions["acot"] = new Acot();
        this.Functions["acoth"] = new Acoth();
        this.Functions["cos"] = new Cos();
        this.Functions["cot"] = new Cot();
        this.Functions["coth"] = new Coth();
        this.Functions["cosh"] = new Cosh();
        this.Functions["csc"] = new Csc();
        this.Functions["csch"] = new Csch();
        this.Functions["power"] = new Power();
        this.Functions["gcd"] = new Gcd();
        this.Functions["lcm"] = new Lcm();
        this.Functions["sec"] = new Sec();
        this.Functions["sech"] = new SecH();
        this.Functions["sign"] = new Sign();
        this.Functions["sqrt"] = new Sqrt();
        this.Functions["sqrtpi"] = new SqrtPi();
        this.Functions["pi"] = new Pi();
        this.Functions["product"] = new Product();
        this.Functions["ceiling"] = new Ceiling();
        this.Functions["ceiling.precise"] = new CeilingPrecise();
        this.Functions["ceiling.math"] = new CeilingMath();
        this.Functions["iso.ceiling"] = new IsoCeiling();
        this.Functions["combin"] = new Combin();
        this.Functions["combina"] = new Combina();
        this.Functions["permut"] = new Permut();
        this.Functions["permutationa"] = new Permutationa();
        this.Functions["count"] = new Count();
        this.Functions["counta"] = new CountA();
        this.Functions["countblank"] = new CountBlank();
        this.Functions["countif"] = new CountIf();
        this.Functions["countifs"] = new CountIfs();
        this.Functions["fact"] = new Fact();
        this.Functions["factdouble"] = new FactDouble();
        this.Functions["floor"] = new Floor();
        this.Functions["floor.precise"] = new FloorPrecise();
        this.Functions["floor.math"] = new FloorMath();
        this.Functions["radians"] = new Radians();
        this.Functions["roman"] = new Roman();
        this.Functions["sin"] = new Sin();
        this.Functions["sinh"] = new Sinh();
        this.Functions["sum"] = new Sum();
        this.Functions["sumif"] = new SumIf();
        this.Functions["sumifs"] = new SumIfs();
        this.Functions["sumproduct"] = new SumProduct();
        this.Functions["sumsq"] = new Sumsq();
        this.Functions["sumxmy2"] = new Sumxmy2();
        this.Functions["sumx2my2"] = new SumX2mY2();
        this.Functions["sumx2py2"] = new SumX2pY2();
        this.Functions["seriessum"] = new Seriessum();
        this.Functions["stdev"] = new Stdev();
        this.Functions["stdeva"] = new Stdeva();
        this.Functions["stdevp"] = new StdevP();
        this.Functions["stdevpa"] = new Stdevpa();
        this.Functions["stdev.s"] = new StdevDotS();
        this.Functions["stdev.p"] = new StdevDotP();
        this.Functions["subtotal"] = new Subtotal();
        this.Functions["exp"] = new Exp();
        this.Functions["log"] = new Log();
        this.Functions["log10"] = new Log10();
        this.Functions["ln"] = new Ln();
        this.Functions["max"] = new Max();
        this.Functions["maxa"] = new Maxa();
        this.Functions["median"] = new Median();
        this.Functions["min"] = new Min();
        this.Functions["mina"] = new Mina();
        this.Functions["mod"] = new Mod();
        this.Functions["mode"] = new Mode();
        this.Functions["mode.sngl"] = new ModeSngl();
        this.Functions["mround"] = new Mround();
        this.Functions["multinomial"] = new Multinomial();
        this.Functions["average"] = new Average();
        this.Functions["averagea"] = new AverageA();
        this.Functions["averageif"] = new AverageIf();
        this.Functions["averageifs"] = new AverageIfs();
        this.Functions["round"] = new Round();
        this.Functions["rounddown"] = new Rounddown();
        this.Functions["roundup"] = new Roundup();
        this.Functions["rand"] = new Rand();
        this.Functions["randbetween"] = new RandBetween();
        this.Functions["rank"] = new Rank();
        this.Functions["rank.eq"] = new RankEq();
        this.Functions["rank.avg"] = new RankAvg();
        this.Functions["percentile"] = new Percentile();
        this.Functions["percentile.inc"] = new PercentileInc();
        this.Functions["percentile.exc"] = new PercentileExc();
        this.Functions["quartile"] = new Quartile();
        this.Functions["quartile.inc"] = new QuartileInc();
        this.Functions["quartile.exc"] = new QuartileExc();
        this.Functions["percentrank"] = new Percentrank();
        this.Functions["percentrank.inc"] = new PercentrankInc();
        this.Functions["percentrank.exc"] = new PercentrankExc();
        this.Functions["quotient"] = new Quotient();
        this.Functions["trunc"] = new Trunc();
        this.Functions["tan"] = new Tan();
        this.Functions["tanh"] = new Tanh();
        this.Functions["atan"] = new Atan();
        this.Functions["atan2"] = new Atan2();
        this.Functions["atanh"] = new Atanh();
        this.Functions["acos"] = new Acos();
        this.Functions["acosh"] = new Acosh();
        this.Functions["covar"] = new Covar();
        this.Functions["covariance.p"] = new CovarianceP();
        this.Functions["covariance.s"] = new CovarianceS();
        this.Functions["var"] = new Var();
        this.Functions["vara"] = new Vara();
        this.Functions["var.s"] = new VarDotS();
        this.Functions["varp"] = new VarP();
        this.Functions["varpa"] = new Varpa();
        this.Functions["var.p"] = new VarDotP();
        this.Functions["large"] = new Large();
        this.Functions["small"] = new Small();
        this.Functions["degrees"] = new Degrees();
        this.Functions["odd"] = new Odd();
        this.Functions["even"] = new Even();

        // Statistical
        this.Functions["confidence.norm"] = new ConfidenceNorm();
        this.Functions["confidence"] = new Confidence();
        this.Functions["confidence.t"] = new ConfidenceT();
        this.Functions["devsq"] = new Devsq();
        this.Functions["avedev"] = new Avedev();
        this.Functions["betadist"] = new Betadist();
        this.Functions["beta.dist"] = new BetaDotDist();
        this.Functions["betainv"] = new Betainv();
        this.Functions["beta.inv"] = new BetaDotInv();
        this.Functions["gamma"] = new Gamma();
        this.Functions["gammaln"] = new Gammaln();
        this.Functions["gammaln.precise"] = new GammalnPrecise();
        this.Functions["norminv"] = new NormInv();
        this.Functions["norm.inv"] = new NormDotInv();
        this.Functions["normsinv"] = new NormsInv();
        this.Functions["norm.s.inv"] = new NormDotSdotInv();
        this.Functions["normdist"] = new Normdist();
        this.Functions["normsdist"] = new Normsdist();
        this.Functions["norm.dist"] = new NormDotDist();
        this.Functions["norm.s.dist"] = new NormDotSdotDist();
        this.Functions["correl"] = new Correl();
        this.Functions["fisher"] = new Fisher();
        this.Functions["fisherinv"] = new FisherInv();
        this.Functions["geomean"] = new Geomean();
        this.Functions["harmean"] = new Harmean();
        this.Functions["pearson"] = new Pearson();
        this.Functions["phi"] = new Phi();
        this.Functions["rsq"] = new Rsq();
        this.Functions["skew"] = new Skew();
        this.Functions["skew.p"] = new SkewP();
        this.Functions["kurt"] = new Kurt();
        this.Functions["gauss"] = new Gauss();
        this.Functions["standardize"] = new Standardize();
        this.Functions["forecast"] = new Forecast();
        this.Functions["forecast.linear"] = new ForecastLinear();
        this.Functions["intercept"] = new Intercept();
        this.Functions["chidist"] = new ChiDist();
        this.Functions["chisq.dist.rt"] = new ChiSqDistRt();
        this.Functions["chisq.inv"] = new ChisqInv();
        this.Functions["chisq.inv.rt"] = new ChisqInvRt();
        this.Functions["chiinv"] = new ChiInv();
        this.Functions["expondist"] = new Expondist();
        this.Functions["expon.dist"] = new ExponDotDist();

        // Information
        this.Functions["isblank"] = new IsBlank();
        this.Functions["isnumber"] = new IsNumber();
        this.Functions["istext"] = new IsText();
        this.Functions["isnontext"] = new IsNonText();
        this.Functions["iserror"] = new IsError();
        this.Functions["iserr"] = new IsErr();
        this.Functions["error.type"] = new ErrorType();
        this.Functions["iseven"] = new IsEven();
        this.Functions["isodd"] = new IsOdd();
        this.Functions["islogical"] = new IsLogical();
        this.Functions["isna"] = new IsNa();
        this.Functions["na"] = new Na();
        this.Functions["n"] = new N();
        this.Functions["type"] = new TypeFunction();
        this.Functions["sheet"] = new Sheet();

        // Logical
        this.Functions["if"] = new If();
        this.Functions["ifs"] = new Ifs();
        this.Functions["maxifs"] = new MaxIfs();
        this.Functions["minifs"] = new MinIfs();
        this.Functions["iferror"] = new IfError();
        this.Functions["ifna"] = new IfNa();
        this.Functions["not"] = new Not();
        this.Functions["and"] = new And();
        this.Functions["or"] = new Or();
        this.Functions["true"] = new True();
        this.Functions["false"] = new False();
        this.Functions["switch"] = new Switch();
        this.Functions["xor"] = new Xor();

        // Reference and lookup
        this.Functions["address"] = new Address();
        this.Functions["hlookup"] = new HLookup();
        this.Functions["vlookup"] = new VLookup();
        this.Functions["lookup"] = new Lookup();
        this.Functions["match"] = new Match();
        this.Functions["row"] = new Row();
        this.Functions["rows"] = new Rows();
        this.Functions["column"] = new Column();
        this.Functions["columns"] = new Columns();
        this.Functions["choose"] = new Choose();
        this.Functions["index"] = new RefAndLookup.Index();
        this.Functions["indirect"] = new Indirect();
        this.Functions["offset"] = new Offset();

        // Date
        this.Functions["date"] = new Date();
        this.Functions["datedif"] = new DateDif();
        this.Functions["today"] = new Today();
        this.Functions["now"] = new Now();
        this.Functions["day"] = new Day();
        this.Functions["month"] = new Month();
        this.Functions["year"] = new Year();
        this.Functions["time"] = new Time();
        this.Functions["hour"] = new Hour();
        this.Functions["minute"] = new Minute();
        this.Functions["second"] = new Second();
        this.Functions["weeknum"] = new Weeknum();
        this.Functions["weekday"] = new Weekday();
        this.Functions["days"] = new Days();
        this.Functions["days360"] = new Days360();
        this.Functions["yearfrac"] = new Yearfrac();
        this.Functions["edate"] = new Edate();
        this.Functions["eomonth"] = new Eomonth();
        this.Functions["isoweeknum"] = new IsoWeekNum();
        this.Functions["workday"] = new Workday();
        this.Functions["workday.intl"] = new WorkdayIntl();
        this.Functions["networkdays"] = new Networkdays();
        this.Functions["networkdays.intl"] = new NetworkdaysIntl();
        this.Functions["datevalue"] = new DateValue();
        this.Functions["timevalue"] = new TimeValue();

        // Database
        this.Functions["dget"] = new Dget();
        this.Functions["dcount"] = new Dcount();
        this.Functions["dcounta"] = new DcountA();
        this.Functions["dmax"] = new Dmax();
        this.Functions["dmin"] = new Dmin();
        this.Functions["dsum"] = new Dsum();
        this.Functions["daverage"] = new Daverage();
        this.Functions["dvar"] = new Dvar();
        this.Functions["dvarp"] = new Dvarp();

        //Finance
        this.Functions["accrint"] = new Accrint();
        this.Functions["accrintm"] = new AccrintM();
        this.Functions["cumipmt"] = new Cumipmt();
        this.Functions["cumprinc"] = new Cumprinc();
        this.Functions["dollarde"] = new DollarDe();
        this.Functions["dollarfr"] = new DollarFr();
        this.Functions["db"] = new Db();
        this.Functions["ddb"] = new Ddb();
        this.Functions["effect"] = new Effect();
        this.Functions["fvschedule"] = new FvSchedule();
        this.Functions["pduration"] = new Pduration();
        this.Functions["rri"] = new Rri();
        this.Functions["pmt"] = new Pmt();
        this.Functions["ppmt"] = new Ppmt();
        this.Functions["ipmt"] = new Ipmt();
        this.Functions["ispmt"] = new IsPmt();
        this.Functions["pv"] = new Pv();
        this.Functions["fv"] = new Fv();
        this.Functions["npv"] = new Npv();
        this.Functions["rate"] = new Rate();
        this.Functions["nper"] = new Nper();
        this.Functions["nominal"] = new Nominal();
        this.Functions["irr"] = new Irr();
        this.Functions["mirr"] = new Mirr();
        this.Functions["xirr"] = new Xirr();
        this.Functions["sln"] = new Sln();
        this.Functions["syd"] = new Syd();
        this.Functions["xnpv"] = new Xnpv();
        this.Functions["coupdays"] = new Coupdays();
        this.Functions["coupdaysnc"] = new Coupdaysnc();
        this.Functions["coupdaybs"] = new Coupdaybs();
        this.Functions["coupnum"] = new Coupnum();
        this.Functions["coupncd"] = new Coupncd();
        this.Functions["couppcd"] = new Couppcd();
        this.Functions["price"] = new Price();
        this.Functions["yield"] = new Yield();
        this.Functions["yieldmat"] = new Yieldmat();
        this.Functions["duration"] = new Duration();
        this.Functions["mduration"] = new Mduration();
        this.Functions["intrate"] = new Intrate();
        this.Functions["disc"] = new Disc();
        this.Functions["tbilleq"] = new Tbilleq();
        this.Functions["tbillprice"] = new TbillPrice();
        this.Functions["tbillyield"] = new TbillYield();

        //Engineering
        this.Functions["bitand"] = new BitAnd();
        this.Functions["bitor"] = new BitOr();
        this.Functions["bitxor"] = new BitXor();
        this.Functions["bitlshift"] = new BitLshift();
        this.Functions["bitrshift"] = new BitRshift();
        this.Functions["convert"] = new ConvertFunction();
        this.Functions["bin2dec"] = new Bin2Dec();
        this.Functions["bin2hex"] = new Bin2Hex();
        this.Functions["bin2oct"] = new Bin2Oct();
        this.Functions["dec2bin"] = new Dec2Bin();
        this.Functions["dec2hex"] = new Dec2Hex();
        this.Functions["dec2oct"] = new Dec2Oct();
        this.Functions["hex2bin"] = new Hex2Bin();
        this.Functions["hex2dec"] = new Hex2Dec();
        this.Functions["hex2oct"] = new Hex2Oct();
        this.Functions["oct2bin"] = new Oct2Bin();
        this.Functions["oct2dec"] = new Oct2Dec();
        this.Functions["oct2hex"] = new Oct2Hex();
        this.Functions["delta"] = new Delta();
        this.Functions["erf"] = new Erf();
        this.Functions["erf.precise"] = new ErfPrecise();
        this.Functions["erfc"] = new Erfc();
        this.Functions["erfc.precise"] = new ErfcPrecise();
        this.Functions["besseli"] = new BesselI();
        this.Functions["besselj"] = new BesselJ();
        this.Functions["besselk"] = new BesselK();
        this.Functions["bessely"] = new BesselY();
        this.Functions["complex"] = new Complex();
    }
}