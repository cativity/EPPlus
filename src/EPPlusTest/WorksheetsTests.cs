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
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Reflection;
using System.Linq;
using System.Runtime.ExceptionServices;

namespace EPPlusTest
{
	[TestClass]
	public class WorksheetsTests
	{
		private ExcelPackage package;
		private ExcelWorkbook workbook;

		[TestInitialize]
		public void TestInitialize()
		{
            this.package = new ExcelPackage();
            this.workbook = this.package.Workbook;
            this.workbook.Worksheets.Add("NEW1");
		}

		[TestMethod]
		public void ConfirmFileStructure()
		{
			Assert.IsNotNull(this.package, "Package not created");
			Assert.IsNotNull(this.workbook, "No workbook found");
		}

		[TestMethod]
		public void ShouldBeAbleToDeleteAndThenAdd()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Delete(1);
            this.workbook.Worksheets.Add("NEW3");
		}

		[TestMethod]
		public void DeleteByNameWhereWorkSheetExists()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Delete("NEW2");
        }

		[TestMethod, ExpectedException(typeof(ArgumentException))]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Delete("NEW3");
		}

		[TestMethod]
		public void MoveBeforeByNameWhereWorkSheetExists()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Add("NEW3");
            this.workbook.Worksheets.Add("NEW4");
            this.workbook.Worksheets.Add("NEW5");

            this.workbook.Worksheets.MoveBefore("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(this.package);
		}

		[TestMethod]
		public void MoveAfterByNameWhereWorkSheetExists()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Add("NEW3");
            this.workbook.Worksheets.Add("NEW4");
            this.workbook.Worksheets.Add("NEW5");

            this.workbook.Worksheets.MoveAfter("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(this.package);
		}

		[TestMethod]
		public void MoveBeforeByPositionWhereWorkSheetExists()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Add("NEW3");
            this.workbook.Worksheets.Add("NEW4");
            this.workbook.Worksheets.Add("NEW5");

            this.workbook.Worksheets.MoveBefore(4, 2);

			CompareOrderOfWorksheetsAfterSaving(this.package);
		}

		[TestMethod]
		public void MoveAfterByPositionWhereWorkSheetExists()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Add("NEW3");
            this.workbook.Worksheets.Add("NEW4");
            this.workbook.Worksheets.Add("NEW5");

            this.workbook.Worksheets.MoveAfter(4, 2);

			CompareOrderOfWorksheetsAfterSaving(this.package);
		}

        [TestMethod]
        public void MoveToStartByNameWhereWorkSheetExists()
        {
            this.workbook.Worksheets.Add("NEW2");

            this.workbook.Worksheets.MoveToStart("NEW2");

            Assert.AreEqual("NEW2", this.workbook.Worksheets.First().Name);
        }

        [TestMethod]
        public void MoveToEndByNameWhereWorkSheetExists()
        {
            this.workbook.Worksheets.Add("NEW2");

            this.workbook.Worksheets.MoveToEnd("NEW1");

            Assert.AreEqual("NEW1", this.workbook.Worksheets.Last().Name);
        }
		[TestMethod]
		public void ShouldHandleResizeOfIndexWhenExceed8Items()
		{
            using ExcelPackage? p = new ExcelPackage();
            ExcelWorksheet wsStart = p.Workbook.Worksheets.Add($"Copy");
            for (int i = 0; i < 7; i++)
            {
                ExcelWorksheet wsNew = p.Workbook.Worksheets.Add($"Sheet{i}");
                p.Workbook.Worksheets.MoveBefore(wsStart.Name, wsNew.Name);
            }
        }
		[TestMethod]
		public void MoveBeforeByName8Worksheets()
		{
            this.workbook.Worksheets.Add("NEW2");
            this.workbook.Worksheets.Add("NEW3");
            this.workbook.Worksheets.Add("NEW4");
            this.workbook.Worksheets.Add("NEW5");
            this.workbook.Worksheets.Add("NEW6");
            this.workbook.Worksheets.Add("NEW7");
            this.workbook.Worksheets.Add("NEW8");

            this.workbook.Worksheets.MoveBefore("NEW8", "NEW1");
			Assert.AreEqual("NEW7", this.workbook.Worksheets.Last().Name);
			Assert.AreEqual("NEW8", this.workbook.Worksheets.First().Name);
			Assert.AreEqual("NEW1", this.workbook.Worksheets[1].Name);
		}
		private static void CompareOrderOfWorksheetsAfterSaving(ExcelPackage editedPackage)
		{
			MemoryStream? packageStream = new MemoryStream();
			editedPackage.SaveAs(packageStream);

			ExcelPackage? newPackage = new ExcelPackage(packageStream);
            int positionId = newPackage._worksheetAdd;
			foreach (ExcelWorksheet? worksheet in editedPackage.Workbook.Worksheets)
			{
				Assert.AreEqual(worksheet.Name, newPackage.Workbook.Worksheets[positionId].Name, "Worksheets are not in the same order");
				positionId++;
			}
		}
        [TestMethod]
        public void CheckAddedWorksheetWithInvalidName()
        {
            if (this.workbook.Worksheets["[NEW2]"] == null)
            {
                this.workbook.Worksheets.Add("[NEW2]");
            }

            Assert.IsNotNull(this.workbook.Worksheets["[NEW2]"]);
        }
    }
}
