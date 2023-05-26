﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ThreadedComments;

namespace EPPlusTest.ThreadedComments
{
    [TestClass]
    public class ThreadedCommentsDeleteTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("ThreadedCommentDelete.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void DeleteOneRowShiftUp()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("OneRowA2");
            ExcelThreadedCommentThread? th=ws.ThreadedComments.Add("A2");
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift up from A2");

            Assert.IsNotNull(ws.Cells["A2"].ThreadedComment);
            ws.DeleteRow(1, 1);
            Assert.IsNull(ws.Cells["A2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteOneColumnShiftLeft()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("OneColumnB1");
            ExcelThreadedCommentThread? th = ws.ThreadedComments.Add("B1");
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift left from B1");

            Assert.IsNotNull(ws.Cells["B1"].ThreadedComment);
            ws.DeleteColumn(1, 1);
            Assert.IsNull(ws.Cells["B1"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteOneRowDeleteThreadedComment()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DeleteA1Row");
            ExcelThreadedCommentThread? th = ws.ThreadedComments.Add("A1");
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "DELTETED!");

            Assert.AreEqual(1, ws.ThreadedComments.Count);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.DeleteRow(1, 1);
            Assert.AreEqual(0, ws.ThreadedComments.Count);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteOneColumnThreadedComment()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("DeleteA1Column");
            ExcelThreadedCommentThread? th = ws.ThreadedComments.Add("A1");
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "DELTETED!");

            Assert.AreEqual(1, ws.ThreadedComments.Count);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.DeleteColumn(1, 1);
            Assert.AreEqual(0, ws.ThreadedComments.Count);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteTwoRowA3()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("A1_A2RowC1");
            ExcelThreadedCommentThread? th = ws.Cells["A3"].AddThreadedComment();
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift down from A1");

            Assert.IsNotNull(ws.Cells["A3"].ThreadedComment);
            ws.Cells["A1:A2"].Delete(eShiftTypeDelete.Up);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNull(ws.Cells["A3"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteTwoColumnC1()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("A1_B1ColumnC1");
            ExcelThreadedCommentThread? th = ws.Cells["C1"].AddThreadedComment();
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift right from A1");

            Assert.IsNotNull(ws.Cells["C1"].ThreadedComment);
            ws.Cells["A1:B1"].Delete(eShiftTypeDelete.Left);
            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNull(ws.Cells["C1"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteInRangeColumn()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("ColumnInRange");
            ExcelThreadedCommentThread? th = ws.Cells["B2:B4"].AddThreadedComment();
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Deleted");
            ws.ThreadedComments["B3"].AddComment(p.Id, "No shift from B3");
            ws.Cells["B4"].ThreadedComment.AddComment(p.Id, "No shift from B4");

            Assert.IsNotNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B3"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B4"].ThreadedComment);
            Assert.AreEqual(3, ws.ThreadedComments.Count);
            ws.Cells["A2:B2"].Delete(eShiftTypeDelete.Left);
            Assert.AreEqual(2, ws.ThreadedComments.Count);
            Assert.IsNotNull(ws.Cells["B3"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B4"].ThreadedComment);
        }
        [TestMethod]
        public void DeleteInRangeRow()
        {
            ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("RowInRange");
            ExcelThreadedCommentThread? th = ws.Cells["B2:D2"].AddThreadedComment();
            ExcelThreadedCommentPerson? p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift down from B2");
            ws.ThreadedComments["C2"].AddComment(p.Id, "No shift from C2");
            ws.Cells["D2"].ThreadedComment.AddComment(p.Id, "No shift from D2");

            Assert.IsNotNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["C2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["D2"].ThreadedComment);
            ws.Cells["B1"].Delete(eShiftTypeDelete.Up);

            Assert.IsNotNull(ws.Cells["B1"].ThreadedComment);
            Assert.IsNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["C2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["D2"].ThreadedComment);
        }
    }
}
