using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ThreadedComments;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ThreadedComments;

[TestClass]
public class PersonsIntegrationTests : TestBase
{
    [TestMethod]
    public void PersonCollOnWorkbook()
    {
        using ExcelPackage? package = OpenTemplatePackage("comments.xlsx");
        ExcelThreadedCommentPersonCollection? persons = package.Workbook.ThreadedCommentPersons;
        _ = persons.Add("Jan Källman", "Jan Källman", IdentityProvider.NoProvider);
        SaveWorkbook("commentsResult.xlsx", package);
    }

    [TestMethod]
    public void AddPersonToWorkbook()
    {
        using ExcelPackage? package = OpenPackage("commentsWithNewPerson.xlsx", true);
        _ = package.Workbook.Worksheets.Add("test");
        ExcelThreadedCommentPersonCollection? persons = package.Workbook.ThreadedCommentPersons;
        _ = persons.Add("Jan Källman", "Jan Källman", IdentityProvider.NoProvider);
        package.Save();
    }
}