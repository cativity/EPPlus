using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class MatchTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            this._parsingContext = ParsingContext.Create();
            this._package = new ExcelPackage();
            this._worksheet = this._package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            this._package.Dispose();
        }

        [TestMethod]
        public void Match_Without_Wildcard()
        {
            this._worksheet.Cells["A1"].Value = "test";
            this._worksheet.Cells["A2"].Value = "value_to_match";
            this._worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            this._worksheet.Cells["A4"].Formula = "MATCH(\"value_to_match\", A1:A3, 0)";

            this._worksheet.Calculate();

            Assert.AreEqual(2, this._worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_Wildcard1()
        {
            this._worksheet.Cells["A1"].Value = "test";
            this._worksheet.Cells["A2"].Value = "value_to_match";
            this._worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            this._worksheet.Cells["A4"].Formula = "MATCH(\"valu*_to_match\", A1:A3, 0)";

            this._worksheet.Calculate();

            Assert.AreEqual(2, this._worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_Wildcard2()
        {
            this._worksheet.Cells["A1"].Value = "test";
            this._worksheet.Cells["A2"].Value = "value_to_match";
            this._worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            this._worksheet.Cells["A4"].Formula = "MATCH(\"?alue_to_match\", A1:A3, 0)";

            this._worksheet.Calculate();

            Assert.AreEqual(2, this._worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_Without_ExactMatch()
        {
            this._worksheet.Cells["A1"].Value = "test";
            this._worksheet.Cells["A2"].Value = "value_to_match";
            this._worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            this._worksheet.Cells["A4"].Formula = "MATCH(\"no_match\", A1:A3, 0)";

            this._worksheet.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this._worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void Match_With_ExactMatch_ShouldNotTreatNullValuesAsAMatch()
        {
            this._worksheet.Cells["A1"].Value = "test";
            this._worksheet.Cells["A2"].Value = "value_to_match";
            this._worksheet.Cells["A3"].Value = "test";

            //_worksheet.Cells["A4"].Value 
            this._worksheet.Cells["A4"].Formula = "MATCH(B1, C1:C2, 0)";

            this._worksheet.Calculate();

            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this._worksheet.Cells["A4"].Value);
        }
    }
}
