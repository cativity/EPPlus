using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation;

[TestClass]
public class UidTests : ValidationTestBase
{
    [TestInitialize]
    public void Setup()
    {
        this.SetupTestData();
    }

    [TestCleanup]
    public void Cleanup()
    {
        this.CleanupTestData();
        this._dataValidationNode = null;
    }

    [TestMethod]
    public void UidShouldBeSetOnValidations()
    {
        // Arrange
        this.LoadXmlTestData("A1", "decimal", "1.3");
        string? id = ExcelDataValidation.NewId();
        // Act
        ExcelDataValidationDecimal? validation = new ExcelDataValidationDecimal(id, "A1", this._sheet);
        // Assert
        Assert.AreEqual(id, validation.Uid);
    }
}