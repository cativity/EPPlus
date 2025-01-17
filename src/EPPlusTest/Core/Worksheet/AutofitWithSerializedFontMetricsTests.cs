﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.SystemDrawing.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet;

[TestClass]
public class AutofitWithSerializedFontMetricsTests : TestBase
{
    [DataTestMethod]
    [DataRow("Calibri")]
    [DataRow("Arial")]
    [DataRow("Arial Black")]
    [DataRow("Times New Roman")]
    [DataRow("Courier New")]
    [DataRow("Liberation Serif")]
    [DataRow("Verdana")]
    [DataRow("Cambria")]
    [DataRow("Cambria Math")]
    [DataRow("Georgia")]
    [DataRow("Corbel")]
    [DataRow("Century Gothic")]
    [DataRow("Rockwell")]
    [DataRow("Trebuchet MS")]
    [DataRow("Tw Cen MT")]
    [DataRow("Tw Cen MT Condensed")]
    public void AutofitWithSerializedFonts(string fontFamily)
    {
        using ExcelPackage? package = new ExcelPackage();

        for (FontSubFamilies style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
        {
            ExcelWorksheet? sheet = package.Workbook.Worksheets.Add(style.ToString());
            ExcelRange? range = sheet.Cells[1, 1, 5, 10];
            range.Style.Font.Name = fontFamily;
            range.Style.Font.Size = 9f;
            range.Style.Font.Italic = style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic;
            range.Style.Font.Bold = style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic;
            Random? rnd = new Random();

            for (int col = 1; col < 10; col++)
            {
                for (int row = 1; row < 5; row++)
                {
                    StringBuilder? sb = new StringBuilder();
                    int maxLength = 40 - (col * 2);
                    int nLetters = rnd.Next(4, maxLength);

                    for (int x = 0; x < nLetters; x++)
                    {
                        int n;
                        if (x % 2 == 0)
                        {
                            n = rnd.Next(65, 90);
                        }
                        else if (x % 5 == 0)
                        {
                            int[]? charArr = new int[] { (int)'.', (int)',', (int)'!', (int)'-' };
                            int cix = rnd.Next(0, charArr.Length - 1);
                            n = charArr[cix];
                        }
                        else if (x % 7 == 0)
                        {
                            n = (int)' ';
                        }
                        else
                        {
                            n = rnd.Next(97, 122);
                        }

                        _ = sb.Append((char)n);
                    }

                    sheet.Cells[row, col].Value = sb.ToString();
                }
            }

            Stopwatch? sw = new Stopwatch();
            sw.Start();
            sheet.Columns[1, 9].AutoFit();
            sw.Stop();
        }

        SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
    }

    //[DataTestMethod, Ignore]

    ////[DataRow("Calibri", 1)]
    ////[DataRow("Calibri Light", 2)]
    ////[DataRow("Arial", 3)]
    ////[DataRow("Arial Black", 4)]
    ////[DataRow("Arial Narrow", 5)]
    ////[DataRow("Bookman Old Style", 6)]
    ////[DataRow("Calisto MT", 7)]
    ////[DataRow("Times New Roman", 8)]
    //[DataRow("Courier New", 9)]

    ////[DataRow("Liberation Serif", 10)]
    ////[DataRow("Verdana", 11)]
    ////[DataRow("Cambria", 12)]
    ////[DataRow("Georgia", 13)]
    ////[DataRow("Corbel", 14)]
    ////[DataRow("Garamond", 15)]
    ////[DataRow("Gill Sans MT", 16)]
    ////[DataRow("Impact", 17)]
    ////[DataRow("Century Gothic", 18)]
    ////[DataRow("Century Schoolbook", 19)]
    ////[DataRow("Rockwell", 20)]
    ////[DataRow("Rockwell Condensed", 21)]
    ////[DataRow("Trebuchet MS", 22)]
    ////[DataRow("Tw Cen MT", 23)]
    ////[DataRow("Tw Cen MT Condensed", 24)]
    //public void AutofitWithSerializedFonts2(string fontFamily, int run)
    //{
    //    ExcelPackage? report = new ExcelPackage(@"c:\Temp\fontreport2.xlsx");
    //    ExcelWorksheet? reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
    //    int reportColOffset = 3;
    //    int reportRow = ((run - 1) * 5) + 2;
    //    List<string>? shortList = new List<string> { "One", "12,3456", "Hello" };
    //    List<string>? mediumList = new List<string> { "A little longer", "5435.1234556", "Something else" };

    //    List<string>? longList = new List<string>
    //    {
    //        "A little longer than the previous example", "5435.1234556", "Something else that is even longer 12345567 than above"
    //    };

    //    List<string>? reallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "Something else that is even longer 12345567 than above, 136542.5439587432 (really, really long)"
    //    };

    //    List<string>? reallyReallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "Something else that is even longer 12345567 than above, 136542.5439587432 (really, really long),,,,,,,,,,,.............&%¤#/%¤)%(/#/%#(%/&¤#`??.3123212321"
    //    };

    //    List<List<string>>? lists = new List<List<string>> { shortList, mediumList, longList, reallyLongList, reallyReallyLongList };
    //    using ExcelPackage? package = new ExcelPackage();
    //    package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
    //    bool newFont = true;

    //    for (FontSubFamilies style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
    //    {
    //        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add(style.ToString());
    //        ExcelRange? range = sheet.Cells[1, 1, 5, 10];
    //        range.Style.Font.Name = fontFamily;
    //        range.Style.Font.Size = 9f;
    //        range.Style.Font.Italic = style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic;
    //        range.Style.Font.Bold = style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic;
    //        //_ = new Random();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            for (int row = 1; row < 4; row++)
    //            {
    //                string? s = lists[col - 1][row - 1];
    //                sheet.Cells[row, col].Value = s;
    //            }
    //        }

    //        Stopwatch? sw = new Stopwatch();
    //        sw.Start();
    //        sheet.Columns[1, 9].AutoFit();

    //        if (newFont)
    //        {
    //            reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
    //            newFont = false;
    //        }

    //        reportSheet.Cells[reportRow, 2].Value = style.ToString();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
    //        }

    //        reportRow++;
    //        sw.Stop();
    //        long ms = sw.ElapsedMilliseconds;
    //    }

    //    SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
    //    report.Save();
    //    report.Dispose();
    //}

    //[DataTestMethod, Ignore]
    //[DataRow("Calibri", 1)]
    //[DataRow("Arial", 2)]
    //[DataRow("Arial Black", 3)]
    //[DataRow("Times New Roman", 4)]
    //[DataRow("Courier New", 5)]
    //[DataRow("Liberation Serif", 6)]
    //[DataRow("Verdana", 7)]
    //[DataRow("Cambria", 8)]
    //[DataRow("Cambria Math", 9)]
    //[DataRow("Georgia", 10)]
    //[DataRow("Corbel", 11)]
    //[DataRow("Century Gothic", 12)]
    //[DataRow("Rockwell", 13)]
    //[DataRow("Trebuchet MS", 14)]
    //[DataRow("Tw Cen MT", 15)]
    //[DataRow("Tw Cen MT Condensed", 16)]
    //[DataRow("MS Gothic", 17)]
    //public void AutofitWithSerializedFonts_JP(string fontFamily, int run)
    //{
    //    ExcelPackage? report = new ExcelPackage(@"c:\Temp\fontreport_jp.xlsx");
    //    ExcelWorksheet? reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
    //    int reportColOffset = 3;
    //    int reportRow = ((run - 1) * 5) + 2;
    //    List<string>? shortList = new List<string> { "新しい最新スタイルです", "ルの拡張サポート", "ピボット テー" };
    //    List<string>? mediumList = new List<string> { "数式計算エンジンの改良点とサポートされる新しい関数", "5435.1234556", "Something else" };
    //    List<string>? longList = new List<string> { "A little longer than the previous example", "5435.1234556", "ェクトが完了すると、コードを管理する開発者のライセンスのみが必要" };

    //    List<string>? reallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "EPPlusは3000万回以上ダウンロードされています。世界中の何千もの企業がスプレッドシートデータを管理するために使用しています。"
    //    };

    //    List<string>? reallyReallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "場合など)、会社は、ユーザーでもあるため、そのサービスの内部ユーザー (開発者) の数をカバーするサブスクリプションをサブスクライブする必要があります。"
    //    };

    //    List<List<string>>? lists = new List<List<string>> { shortList, mediumList, longList, reallyLongList, reallyReallyLongList };
    //    using ExcelPackage? package = new ExcelPackage();
    //    package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
    //    bool newFont = true;

    //    for (FontSubFamilies style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
    //    {
    //        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add(style.ToString());
    //        ExcelRange? range = sheet.Cells[1, 1, 5, 10];
    //        range.Style.Font.Name = fontFamily;
    //        range.Style.Font.Size = 24f;
    //        range.Style.Font.Italic = style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic;
    //        range.Style.Font.Bold = style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic;
    //        //_ = new Random();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            for (int row = 1; row < 4; row++)
    //            {
    //                string? s = lists[col - 1][row - 1];
    //                sheet.Cells[row, col].Value = s;
    //            }
    //        }

    //        Stopwatch? sw = new Stopwatch();
    //        sw.Start();
    //        sheet.Columns[1, 9].AutoFit();

    //        if (newFont)
    //        {
    //            reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
    //            newFont = false;
    //        }

    //        reportSheet.Cells[reportRow, 2].Value = style.ToString();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
    //        }

    //        reportRow++;
    //        sw.Stop();
    //        long ms = sw.ElapsedMilliseconds;
    //    }

    //    SaveWorkbook($"JP_Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
    //    report.Save();
    //    report.Dispose();
    //}

    [TestMethod]
    public void LoadFontSizeFromResource()
    {
        using ExcelPackage? p = new ExcelPackage();
        int expectedLoaded = 895;

        if (FontSize._isLoaded == false)
        {
            int expectedDefault = 23;
            Assert.AreEqual(expectedDefault, FontSize.FontHeights.Count);
            Assert.AreEqual(expectedDefault, FontSize.FontWidths.Count);
        }

        FontSize.LoadAllFontsFromResource();
        Assert.AreEqual(expectedLoaded, FontSize.FontHeights.Count);
        Assert.AreEqual(expectedLoaded, FontSize.FontWidths.Count);
    }

    //[DataTestMethod, Ignore]
    //[DataRow("Calibri")]
    //[DataRow("Arial")]
    //[DataRow("Times New Roman")]
    //public void MeasureSpecificFont(string font)
    //{
    //    using ExcelPackage? package = new ExcelPackage();
    //    package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
    //    ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("text");
    //    ExcelWorksheet? sheet2 = package.Workbook.Worksheets.Add("measures");
    //    ExcelWorksheet? sheet3 = package.Workbook.Worksheets.Add("numbers");
    //    sheet.Cells["A1:A50"].Style.Font.Name = font;
    //    sheet.Cells["A1:A50"].Style.Font.Italic = true;
    //    string? chars = "aabcdeefghijklmnopqrrssttuvxyzåäö   AABCDEEFGHIJKLMNOPQRSSTTUVXYZÅÄÖ      !!,,,,,,,,, 112233445566778899.....";
    //    string? numbers = "11122233344455566677788899900000000........,,,,,,,       ";
    //    Random? rnd = new Random();

    //    for (int x = 0; x < 60; x++)
    //    {
    //        StringBuilder? text = new StringBuilder();

    //        for (int i = 0; i < x; i++)
    //        {
    //            int ix = rnd.Next(0, chars.Length);
    //            text.Append(chars[ix]);
    //        }

    //        sheet.Cells[1, x + 1].Value = text.ToString();
    //        sheet.Columns[x + 1].AutoFit();
    //        sheet2.Cells[1, x + 1].Value = sheet.Columns[x + 1].Width;

    //        StringBuilder? number = new StringBuilder();

    //        for (int i = 0; i < x; i++)
    //        {
    //            int ix = rnd.Next(0, numbers.Length);
    //            number.Append(numbers[ix]);
    //        }

    //        sheet3.Cells[1, x + 1].Value = number.ToString();
    //        sheet3.Columns[x + 1].AutoFit();
    //        sheet2.Cells[2, x + 2].Value = sheet3.Columns[x + 1].Width;
    //    }

    //    if (!Directory.Exists(@"c:\Temp\FontTests"))
    //    {
    //        Directory.CreateDirectory(@"c:\Temp\FontTests");
    //    }

    //    string? path = $"c:\\Temp\\FontTests\\{font}Measurements.xlsx";

    //    if (File.Exists(path))
    //    {
    //        File.Delete(path);
    //    }

    //    package.SaveAs(path);
    //}

    //[DataTestMethod, Ignore]
    //[DataRow("Yu Gothic", 1)]
    //[DataRow("Yu Mincho", 2)]
    //[DataRow("Arial Rounded MT Bold", 3)]
    //[DataRow("Goudy Stout", 4)]
    //[DataRow("Vladimir Script", 5)]
    //[DataRow("Bahnschrift SemiBold SemiConden", 6)]
    //[DataRow("Copperplate Gothic Bold", 7)]
    //[DataRow("Gigi", 8)]
    //[DataRow("Non existing font", 9)]
    //public void MeasureOtherFonts(string fontFamily, int run)
    //{
    //    ExcelPackage? report = new ExcelPackage(@"c:\Temp\fontreport_jp.xlsx");
    //    ExcelWorksheet? reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
    //    int reportColOffset = 3;
    //    int reportRow = ((run - 1) * 5) + 2;
    //    List<string>? shortList = new List<string> { "新しい最新スタイルです", "ルの拡張サポート", "ピボット テー" };
    //    List<string>? mediumList = new List<string> { "数式計算エンジンの改良点とサポートされる新しい関数", "5435.1234556", "Something else" };
    //    List<string>? longList = new List<string> { "A little longer than the previous example", "5435.1234556", "ェクトが完了すると、コードを管理する開発者のライセンスのみが必要" };

    //    List<string>? reallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "EPPlusは3000万回以上ダウンロードされています。世界中の何千もの企業がスプレッドシートデータを管理するために使用しています。"
    //    };

    //    List<string>? reallyReallyLongList = new List<string>
    //    {
    //        "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
    //        "5435.1234556321 - 4.32413254353",
    //        "場合など)、会社は、ユーザーでもあるため、そのサービスの内部ユーザー (開発者) の数をカバーするサブスクリプションをサブスクライブする必要があります。"
    //    };

    //    List<List<string>>? lists = new List<List<string>> { shortList, mediumList, longList, reallyLongList, reallyReallyLongList };
    //    using ExcelPackage? package = new ExcelPackage();

    //    //package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
    //    bool newFont = true;

    //    for (FontSubFamilies style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
    //    {
    //        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add(style.ToString());
    //        ExcelRange? range = sheet.Cells[1, 1, 5, 10];
    //        range.Style.Font.Name = fontFamily;
    //        range.Style.Font.Size = 24f;
    //        range.Style.Font.Italic = style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic;
    //        range.Style.Font.Bold = style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic;
    //        //_ = new Random();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            for (int row = 1; row < 2; row++)
    //            {
    //                string? s = lists[col - 1][row - 1];
    //                sheet.Cells[row, col].Value = s;
    //            }
    //        }

    //        Stopwatch? sw = new Stopwatch();
    //        sw.Start();
    //        sheet.Columns[1, 9].AutoFit();

    //        if (newFont)
    //        {
    //            reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
    //            newFont = false;
    //        }

    //        reportSheet.Cells[reportRow, 2].Value = style.ToString();

    //        for (int col = 1; col < lists.Count + 1; col++)
    //        {
    //            reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
    //        }

    //        reportRow++;
    //        sw.Stop();
    //        long ms = sw.ElapsedMilliseconds;
    //    }

    //    SaveWorkbook($"NonExistingFonts_Autofit_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
    //    report.Save();
    //    report.Dispose();
    //}

//#if NETFULL
//        [TestMethod, Ignore]
//        public void AutoFitSystemDrawing()
//        {
//            using(var package = new ExcelPackage())
//            {
//                //package.Workbook.TextSettings.FallbackTextMeasurer = new OfficeOpenXml.SkiaSharp.Text.SkiaSharpTextMeasurer();
//                //var sheet = package.Workbook.Worksheets.Add("Test");
//                //sheet.Cells["A1"].Value = "abc 123 SDFÖLKJE !wueriopiquwejklöpasdfj";
//                //sheet.Cells["A1"].Style.Font.Name = "Times New Roman";
//                //sheet.Columns.AutoFit();
//                //SaveWorkbook("Autofit_Candara.xlsx", package);
//            }
//        }
//#endif
}