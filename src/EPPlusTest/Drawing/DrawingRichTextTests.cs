﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.IO;
using OfficeOpenXml.Style;

namespace EPPlusTest.Drawing
{
    namespace EPPlusTest.Drawing
    {
        [TestClass]
        public class DrawingRichTextTests : TestBase
        {
            static ExcelPackage _pck;
            static ExcelWorksheet _ws;
            [ClassInitialize]
            public static void Init(TestContext context)
            {
                _pck = OpenPackage("DrawingRichText.xlsx", true);                
                _ws = _pck.Workbook.Worksheets.Add("Richtext");
            }
            [ClassCleanup]
            public static void Cleanup()
            {
                string? dirName = _pck.File.DirectoryName;
                string? fileName = _pck.File.FullName;

                SaveAndCleanup(_pck);

                File.Copy(fileName, dirName + "\\DrawingRichTextRead.xlsx", true);
            }

            [TestMethod]
            public void AddThreeParagraphsAndValidate()
            {
                ExcelShape? shape = _ws.Drawings.AddShape("shape1", eShapeStyle.Rect);
                shape.RichText.Add("Line1");
                ExcelParagraph? r2=shape.RichText.Add("L", true);
                r2.Fill.Style = eFillStyle.SolidFill;
                r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent2);
                r2 = shape.RichText.Add("i");
                r2.Fill.Style = eFillStyle.SolidFill;
                r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent3);
                r2 = shape.RichText.Add("n");
                r2.Fill.Style = eFillStyle.SolidFill;
                r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent4);
                r2 = shape.RichText.Add("e");
                r2.Fill.Style = eFillStyle.SolidFill;
                r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent5);
                r2 = shape.RichText.Add("2");
                r2.Fill.Style = eFillStyle.SolidFill;
                r2.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);


                ExcelParagraph? r3=shape.RichText.Add("Line3", true);
                r3.Bold = true;
                r3.Italic = true;
                r3.LatinFont = "Times New Roman";
                r3.Size = 19.5F;

                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.Text);
                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.RichText.Text);

                Assert.AreEqual(7, shape.RichText.Count);
                Assert.IsTrue(shape.RichText[0].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[0].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[1].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[1].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[2].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[2].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[3].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[3].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[4].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[4].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[5].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[5].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[6].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[6].IsLastInParagraph);
            }
            [TestMethod]
            public void ReadThreeParagraphsAndValidate()
            {
                AssertIfNotExists("DrawingRichTextRead.xlsx");
                using ExcelPackage? p = OpenPackage("DrawingRichTextRead.xlsx");
                ExcelShape? shape = (ExcelShape)p.Workbook.Worksheets[0].Drawings["shape1"];
                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.Text);
                Assert.AreEqual("Line1\r\nLine2\r\nLine3", shape.RichText.Text);

                Assert.AreEqual(7, shape.RichText.Count);
                Assert.IsTrue(shape.RichText[0].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[0].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[1].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[1].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[2].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[2].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[3].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[3].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[4].IsFirstInParagraph);
                Assert.IsFalse(shape.RichText[4].IsLastInParagraph);
                Assert.IsFalse(shape.RichText[5].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[5].IsLastInParagraph);
                Assert.IsTrue(shape.RichText[6].IsFirstInParagraph);
                Assert.IsTrue(shape.RichText[6].IsLastInParagraph);
            }
            [TestMethod]
            public void AddEmptyParagraphFirst()
            {
                ExcelShape? shape = _ws.Drawings.AddShape("shape2", eShapeStyle.Rect);
                shape.SetPosition(20, 0, 0, 0);
                shape.RichText.Add("", true);
                shape.RichText.Add("SecondLine", true);
                ExcelParagraph? r2 = shape.RichText.Add("    ", true);
                r2.UnderLine = eUnderLineType.Single;
                Assert.AreEqual(3, shape.RichText.Count);
                Assert.AreEqual("", shape.RichText[0].Text);
                Assert.AreEqual("SecondLine", shape.RichText[1].Text);
                Assert.AreEqual("    ", shape.RichText[2].Text);
            }
        }
    }
}
