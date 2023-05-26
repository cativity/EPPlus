/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class DateTimeFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        private double GetTime(int hour, int minute, int second)
        {
            double secInADay = DateTime.Today.AddDays(1).Subtract(DateTime.Today).TotalSeconds;
            double secondsOfExample = (double)(hour * 60 * 60 + minute * 60 + second);
            return secondsOfExample / secInADay;
        }
        [TestMethod]
        public void DateFunctionShouldReturnADate()
        {
            Date? func = new Date();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2012, 4, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.Date, result.DataType);
        }

        [TestMethod]
        public void DateFunctionShouldReturnACorrectDate()
        {
            DateTime expectedDate = new DateTime(2012, 4, 3);
            Date? func = new Date();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2012, 4, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
        {
            DateTime expectedDate = new DateTime(2011, 11, 3);
            Date? func = new Date();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(2012, -1, 3);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void NowFunctionShouldReturnNow()
        {
            DateTime startTime = DateTime.Now;
            Thread.Sleep(1);
            Now? func = new Now();
            FunctionArgument[]? args = new FunctionArgument[0];
            CompileResult? result = func.Execute(args, _parsingContext);
            Thread.Sleep(1);
            DateTime endTime = DateTime.Now;
            DateTime resultDate = DateTime.FromOADate((double)result.Result);
            Assert.IsTrue(resultDate > startTime && resultDate < endTime);
        }

        [TestMethod]
        public void TodayFunctionShouldReturnTodaysDate()
        {
            Today? func = new Today();
            FunctionArgument[]? args = new FunctionArgument[0];
            CompileResult? result = func.Execute(args, _parsingContext);
            DateTime resultDate = DateTime.FromOADate((double)result.Result);
            Assert.AreEqual(DateTime.Now.Date, resultDate);
        }

        [TestMethod]
        public void DayShouldReturnDayInMonth()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Day? func = new Day();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(date.ToOADate());
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void DayShouldReturnMonthOfYearWithStringParam()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Day? func = new Day();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void DaysShouldReturnCorrectResultWithDateTimeTypes()
        {
            DateTime d1 = new DateTime(2015, 1, 1);
            DateTime d2 = new DateTime(2015, 2, 2);
            Days? func = new Days();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(d2, d1), _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void DaysShouldReturnCorrectResultWithDatesAsStrings()
        {
            string? d1 = new DateTime(2015, 1, 1).ToString("yyyy-MM-dd");
            string? d2 = new DateTime(2015, 2, 2).ToString("yyyy-MM-dd");
            Days? func = new Days();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(d2, d1), _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void MonthShouldReturnMonthOfYear()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Month? func = new Month();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void MonthShouldReturnMonthOfYearWithStringParam()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Month? func = new Month();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectYear()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Year? func = new Year();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
            Assert.AreEqual(2012, result.Result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectYearWithStringParam()
        {
            DateTime date = new DateTime(2012, 3, 12);
            Year? func = new Year();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(2012, result.Result);
        }

        [TestMethod]
        public void TimeShouldReturnACorrectSerialNumber()
        {
            double expectedResult = GetTime(10, 11, 12);
            Time? func = new Time();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(10, 11, 12), _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);  
        }

        [TestMethod]
        public void TimeShouldParseStringCorrectly()
        {
            double expectedResult = GetTime(10, 11, 12);
            Time? func = new Time();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("10:11:12"), _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void TimeShouldReturnErrorIsOutOfRange()
        {
            Time? func = new Time();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(10, 11, 60), _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void TimeShouldReturnErrorIfMinuteIsOutOfRange()
        {
            Time? func = new Time();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(10, 60, 12), _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void TimeShouldReturnErrorIfHourIsOutOfRange()
        {
            Time? func = new Time();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(24, 12, 12), _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            Hour? func = new Hour();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 13, 14)), _parsingContext);
            Assert.AreEqual(9, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(GetTime(23, 13, 14)), _parsingContext);
            Assert.AreEqual(23, result.Result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            Minute? func = new Minute();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 14)), _parsingContext);
            Assert.AreEqual(14, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 55, 14)), _parsingContext);
            Assert.AreEqual(55, result.Result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            Second? func = new Second();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 17)), _parsingContext);
            Assert.AreEqual(17, result.Result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResultWithStringArgument()
        {
            Second? func = new Second();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResultWithStringArgument()
        {
            Minute? func = new Minute();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(11, result.Result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResultWithStringArgument()
        {
            Hour? func = new Hour();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(10, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs1()
        {
            Weekday? func = new Weekday();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 1), _parsingContext);
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs2()
        {
            Weekday? func = new Weekday();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 2), _parsingContext);
            Assert.AreEqual(7, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs3()
        {
            Weekday? func = new Weekday();
            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 3), _parsingContext);
            Assert.AreEqual(6, result.Result);
        }

        [TestMethod]
        public void WeekNumShouldReturnCorrectResult()
        {
            Weeknum? func = new Weeknum();
            double dt1 = new DateTime(2012, 12, 31).ToOADate();
            double dt2 = new DateTime(2012, 1, 1).ToOADate();
            double dt3 = new DateTime(2013, 1, 20).ToOADate();

            CompileResult? r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), _parsingContext);
            CompileResult? r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), _parsingContext);
            CompileResult? r3 = func.Execute(FunctionsHelper.CreateArgs(dt3, 2), _parsingContext);

            Assert.AreEqual(53, r1.Result, "r1.Result was not 53, but " + r1.Result.ToString());
            Assert.AreEqual(1, r2.Result, "r2.Result was not 1, but " + r2.Result.ToString());
            Assert.AreEqual(3, r3.Result, "r3.Result was not 3, but " + r3.Result.ToString());
        }

        [TestMethod]
        public void EdateShouldReturnCorrectResult()
        {
            Edate? func = new Edate();

            double dt1arg = new DateTime(2012, 1, 31).ToOADate();
            double dt2arg = new DateTime(2013, 1, 1).ToOADate();
            double dt3arg = new DateTime(2013, 2, 28).ToOADate();

            CompileResult? r1 = func.Execute(FunctionsHelper.CreateArgs(dt1arg, 1), _parsingContext);
            CompileResult? r2 = func.Execute(FunctionsHelper.CreateArgs(dt2arg, -1), _parsingContext);
            CompileResult? r3 = func.Execute(FunctionsHelper.CreateArgs(dt3arg, 2), _parsingContext);

            DateTime dt1 = DateTime.FromOADate((double) r1.Result);
            DateTime dt2 = DateTime.FromOADate((double)r2.Result);
            DateTime dt3 = DateTime.FromOADate((double)r3.Result);

            DateTime exp1 = new DateTime(2012, 2, 29);
            DateTime exp2 = new DateTime(2012, 12, 1);
            DateTime exp3 = new DateTime(2013, 4, 28);

            Assert.AreEqual(exp1, dt1, "dt1 was not " + exp1.ToString("yyyy-MM-dd") + ", but " + dt1.ToString("yyyy-MM-dd"));
            Assert.AreEqual(exp2, dt2, "dt1 was not " + exp2.ToString("yyyy-MM-dd") + ", but " + dt2.ToString("yyyy-MM-dd"));
            Assert.AreEqual(exp3, dt3, "dt1 was not " + exp3.ToString("yyyy-MM-dd") + ", but " + dt3.ToString("yyyy-MM-dd"));
        }

        [TestMethod]
        public void Days360ShouldReturnCorrectResultWithNoMethodSpecified2()
        {
            Days360? func = new Days360();

            double dt1arg = new DateTime(2013, 1, 1).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), _parsingContext);

            Assert.AreEqual(90, result.Result);
        }

        [TestMethod]
        public void Days360ShouldReturnCorrectResultWithEuroMethodSpecified()
        {
            Days360? func = new Days360();

            double dt1arg = new DateTime(2013, 1, 1).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

            Assert.AreEqual(89, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithEuroMethodSpecified()
        {
            Days360? func = new Days360();

            double dt1arg = new DateTime(2012, 2, 29).ToOADate();
            double dt2arg = new DateTime(2013, 2, 28).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

            Assert.AreEqual(359, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithUsMethodSpecified()
        {
            Days360? func = new Days360();

            double dt1arg = new DateTime(2012, 2, 29).ToOADate();
            double dt2arg = new DateTime(2013, 2, 28).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

            Assert.AreEqual(358, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithUsMethodSpecified2()
        {
            Days360? func = new Days360();

            double dt1arg = new DateTime(2013, 2, 28).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

            Assert.AreEqual(30, result.Result);
        }

        [TestMethod]
        public void YearFracShouldReturnCorrectResultWithUsBasis()
        {
            Yearfrac? func = new Yearfrac();
            double dt1arg = new DateTime(2013, 2, 28).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), _parsingContext);

            double roundedResult = Math.Round((double) result.Result, 4);

            Assert.IsTrue(Math.Abs(0.0861 - roundedResult) < double.Epsilon);
        }

        [TestMethod]
        public void YearFracShouldReturnCorrectResultWithEuroBasis()
        {
            Yearfrac? func = new Yearfrac();
            double dt1arg = new DateTime(2013, 2, 28).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, 4), _parsingContext);

            double roundedResult = Math.Round((double)result.Result, 4);

            Assert.IsTrue(Math.Abs(0.0889 - roundedResult) < double.Epsilon);
        }

        [TestMethod]
        public void YearFracActualActual()
        {
            Yearfrac? func = new Yearfrac();
            double dt1arg = new DateTime(2012, 2, 28).ToOADate();
            double dt2arg = new DateTime(2013, 3, 31).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, 1), _parsingContext);

            double roundedResult = Math.Round((double)result.Result, 4);

            Assert.IsTrue(Math.Abs(1.0862 - roundedResult) < double.Epsilon);
        }

        [TestMethod]
        public void IsoWeekShouldReturn1When1StJan()
        {
            IsoWeekNum? func = new IsoWeekNum();
            double arg = new DateTime(2013, 1, 1).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(arg), _parsingContext);

            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void EomonthShouldReturnCorrectResultWithPositiveArg()
        {
            Eomonth? func = new Eomonth();
            double arg = new DateTime(2013, 2, 2).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(arg, 3), _parsingContext);

            Assert.AreEqual(41425d, result.Result);
        }

        [TestMethod]
        public void EomonthShouldReturnCorrectResultWithNegativeArg()
        {
            Eomonth? func = new Eomonth();
            double arg = new DateTime(2013, 2, 2).ToOADate();

            CompileResult? result = func.Execute(FunctionsHelper.CreateArgs(arg, -3), _parsingContext);

            Assert.AreEqual(41243d, result.Result);
        }

        [TestMethod]
        public void WorkdayShouldReturnCorrectResultIfNoHolidayIsSupplied()
        {
            double inputDate = new DateTime(2014, 1, 1).ToOADate();
            double expectedDate = new DateTime(2014, 1, 29).ToOADate();

            Workday? func = new Workday();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(inputDate, 20);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate, result.Result);
        }

        [TestMethod]
        public void WorkdayShouldReturnCorrectResultWithNegativeArg()
        {
            double inputDate = new DateTime(2016, 6, 15).ToOADate();
            double expectedDate = new DateTime(2016, 5, 4).ToOADate();

            Workday? func = new Workday();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(inputDate, -30);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DateTime.FromOADate(expectedDate), DateTime.FromOADate((double)result.Result));
        }

        [TestMethod]
        public void WorkdayShouldReturnCorrectResultWithFourDaysSupplied()
        {
            double inputDate = new DateTime(2014, 1, 1).ToOADate();
            double expectedDate = new DateTime(2014, 1, 7).ToOADate();

            Workday? func = new Workday();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(inputDate, 4);
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate, result.Result);
        }

        [TestMethod]
        public void WorkdayWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
        {
            double inputDate = new DateTime(2016, 7, 27).ToOADate();
            double holidayDate1 = new DateTime(2016, 7, 11).ToOADate();
            double holidayDate2 = new DateTime(2016, 7, 8).ToOADate();
            double expectedDate = new DateTime(2016, 6, 13).ToOADate();

            Workday? func = new Workday();
            IEnumerable<FunctionArgument>? args = FunctionsHelper.CreateArgs(inputDate, -30, FunctionsHelper.CreateArgs(holidayDate1, holidayDate2));
            CompileResult? result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate, result.Result);
        }

        [TestMethod]
        public void WorkdayWithNegativeArgShouldReturnCorrectWhenRangeWithHolidayDatesIsSupplied()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Value = new DateTime(2016, 7, 27).ToOADate();
                ws.Cells["B1"].Value = new DateTime(2016, 7, 11).ToOADate();
                ws.Cells["B2"].Value = new DateTime(2016, 7, 8).ToOADate();
                ws.Cells["B3"].Formula = "WORKDAY(A1,-30, B1:B2)";
                ws.Calculate();

                double expectedDate = new DateTime(2016, 6, 13).ToOADate();
                object? actualDate = ws.Cells["B3"].Value;
                Assert.AreEqual(expectedDate, actualDate);
            } 
        }

        [TestMethod]
        public void WorkdayIntlShouldCalculateWeekendOnly()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["B1"].Value = new DateTime(2015, 12, 1).ToOADate();
                ws.Cells["B3"].Formula = "WORKDAY.INTL(B1,25)";
                ws.Calculate();

                double expectedDate = new DateTime(2016, 1, 5).ToOADate();
                object? actualDate = ws.Cells["B3"].Value;
                Assert.AreEqual(expectedDate, actualDate);
            }
        }

        [TestMethod]
        public void WorkdayIntlShouldCalculateHolidaysStringArg()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["B1"].Value = new DateTime(2015, 12, 1).ToOADate();
                ws.Cells["B2"].Value = new DateTime(2015, 12, 25).ToOADate();
                ws.Cells["B3"].Value = new DateTime(2015, 12, 28).ToOADate();
                ws.Cells["B4"].Value = new DateTime(2016, 1, 1).ToOADate();
                ws.Cells["B7"].Formula = "WORKDAY.INTL(B1,25,\"0000111\")";
                ws.Calculate();

                double expectedDate = new DateTime(2016, 1, 13).ToOADate();
                object? actualDate = ws.Cells["B7"].Value;
                Assert.AreEqual(expectedDate, actualDate);
            }
        }

        [TestMethod]
        public void WorkdayIntlShouldCalculateHolidays()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["B1"].Value = new DateTime(2015, 12, 1).ToOADate();
                ws.Cells["B2"].Value = new DateTime(2015, 12, 25).ToOADate();
                ws.Cells["B3"].Value = new DateTime(2015, 12, 28).ToOADate();
                ws.Cells["B4"].Value = new DateTime(2016, 1, 1).ToOADate();
                ws.Cells["B7"].Formula = "WORKDAY.INTL(B1,25,1,B2:B4)";
                ws.Calculate();

                double expectedDate = new DateTime(2016, 1, 8).ToOADate();
                object? actualDate = ws.Cells["B7"].Value;
                Assert.AreEqual(expectedDate, actualDate);
            }
        }

        [TestMethod]
        public void NetworkdaysShouldReturnNumberOfDays()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS(DATE(2016,1,1), DATE(2016,1,20))";
                ws.Calculate();
                Assert.AreEqual(14, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void NetworkdaysShouldReturnNumberOfDaysWithHolidayRange()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // do something...
                using (ExcelPackage? package = new ExcelPackage())
                {
                    package.Load(ms);
                }
            }
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS(DATE(2016,1,1), DATE(2016,1,20),B1)";
                ws.Cells["B1"].Formula = "DATE(2016,1,15)";
                ws.Calculate();
                Assert.AreEqual(13, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void NetworkdaysNegativeShouldReturnNumberOfDays()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS(DATE(2016,1,1), DATE(2015,12,20))";
                ws.Calculate();
                Assert.AreEqual(10, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void NetworkdayIntlShouldUseWeekendArg()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), 11)";
                ws.Calculate();
                Assert.AreEqual(17, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void NetworkdayIntlShouldUseWeekendStringArg()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), \"0000011\")";
                ws.Calculate();
                Assert.AreEqual(14, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void NetworkdayIntlShouldReduceHoliday()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "NETWORKDAYS.INTL(DATE(2016,1,1), DATE(2016,1,20), \"0000011\", DATE(2016,1,4))";
                ws.Calculate();
                Assert.AreEqual(13, ws.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void TimeAddition()
        {
            using (ExcelPackage? package = new ExcelPackage())
            {
                ExcelWorksheet? ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "1 + Time(10,0,0)";
                ws.Calculate();
                double result = Convert.ToDouble(ws.Cells["A1"].Value);
                result = Math.Round(result, 2);
                Assert.AreEqual(1.42d, result);
            }
        }
    }
}
