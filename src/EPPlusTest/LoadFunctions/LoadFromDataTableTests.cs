﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.LoadFunctions;

[TestClass]
public class LoadFromDataTableTests
{
    [TestInitialize]
    public void Initialize()
    {
        this._package = new ExcelPackage();
        this._worksheet = this._package.Workbook.Worksheets.Add("test");
        this._dataSet = new DataSet();
        this._table = this._dataSet.Tables.Add("table");
        _ = this._table.Columns.Add("Id", typeof(string));
        _ = this._table.Columns.Add("Name", typeof(string));
    }

    [TestCleanup]
    public void Cleanup() => this._package.Dispose();

    private ExcelPackage _package;
    private ExcelWorksheet _worksheet;
    private DataSet _dataSet;
    private DataTable _table;

    [TestMethod]
    public void ShouldLoadTable()
    {
        _ = this._table.Rows.Add("1", "Test name");
        _ = this._worksheet.Cells["A1"].LoadFromDataTable(this._table, false);
        Assert.AreEqual("1", this._worksheet.Cells["A1"].Value);
    }

    [TestMethod]
    public void CreateAndFillDataTable()
    {
        DataTable? table = new DataTable("Astronauts");
        _ = table.Columns.Add("Id", typeof(int));
        _ = table.Columns.Add("FirstName", typeof(string));
        _ = table.Columns.Add("LastName", typeof(string));
        table.Columns["FirstName"].Caption = "First name";
        table.Columns["LastName"].Caption = "Last name";

        // add some data
        _ = table.Rows.Add(1, "Bob", "Behnken");
        _ = table.Rows.Add(2, "Doug", "Hurley");

        //create a workbook with a spreadsheet and load the data table
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("Astronauts");
        _ = sheet.Cells["A1"].LoadFromDataTable(table);
    }

    [TestMethod]
    public void ShouldLoadTableWithTableStyle()
    {
        _ = this._table.Rows.Add("1", "Test name");
        _ = this._worksheet.Cells["A1"].LoadFromDataTable(this._table, false, TableStyles.Dark1);
        Assert.AreEqual(1, this._worksheet.Tables.Count);
    }

    [TestMethod]
    public void ShouldUseCaptionForHeader()
    {
        this._table.Columns["Id"].Caption = "An Id";
        this._table.Columns["Name"].Caption = "A name";
        _ = this._worksheet.Cells["A1"].LoadFromDataTable(this._table, true);
        Assert.AreEqual("An Id", this._worksheet.Cells["A1"].Value);
    }

    [TestMethod]
    public void ShouldUseColumnNameForHeaderIfNoCaption()
    {
        _ = this._worksheet.Cells["A1"].LoadFromDataTable(this._table, true);
        Assert.AreEqual("Id", this._worksheet.Cells["A1"].Value);
    }

    [TestMethod]
    public void ShouldLoadXmlFromDataset()
    {
        DataSet? dataSet = new DataSet();

        string? xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                      + "<Astronauts>"
                      + "<Astronaut Id=\"1\">"
                      + "<FirstName>Bob</FirstName>"
                      + "<LastName>Behnken</LastName>"
                      + "</Astronaut>"
                      + "<Astronaut Id=\"2\">"
                      + "<FirstName>Doug</FirstName>"
                      + "<LastName>Hurley</LastName>"
                      + "</Astronaut>"
                      + "</Astronauts>";

        XmlReader? reader = XmlReader.Create(new StringReader(xml));
        _ = dataSet.ReadXml(reader);
        using ExcelPackage? package = new ExcelPackage();
        ExcelWorksheet? sheet = package.Workbook.Worksheets.Add("test");
        DataTable? table = dataSet.Tables["Astronaut"];

        // default the Id ends up last in the column order. This moves it to the first position.
        table.Columns["Id"].SetOrdinal(0);

        // Set caption for the headers
        table.Columns["FirstName"].Caption = "First name";
        table.Columns["LastName"].Caption = "Last name";

        // call LoadFromDataTable, print headers and use the Dark1 table style
        _ = sheet.Cells["A1"].LoadFromDataTable(table, true, TableStyles.Dark1);

        // AutoFit column with for the entire range
        sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Row].AutoFitColumns();

        //package.SaveAs(new FileInfo(@"c:\temp\astronauts.xlsx"));
    }

    [TestMethod]
    public void ShouldUseLambdaConfig()
    {
        _ = this._table.Rows.Add("1", "Test name");

        _ = this._worksheet.Cells["A1"]
            .LoadFromDataTable(this._table,
                               c =>
                               {
                                   c.PrintHeaders = true;
                                   c.TableStyle = TableStyles.Dark1;
                               });

        Assert.AreEqual("Id", this._worksheet.Cells["A1"].Value);
        Assert.AreEqual("1", this._worksheet.Cells["A2"].Value);
    }

    [TestMethod]
    public void ShouldSetDbNullToNull()
    {
        _ = this._table.Rows.Add("1", DBNull.Value);

        _ = this._worksheet.Cells["A1"]
            .LoadFromDataTable(this._table,
                               c =>
                               {
                                   c.PrintHeaders = true;
                                   c.TableStyle = TableStyles.Dark1;
                               });

        Assert.IsNull(this._worksheet.Cells["B2"].Value);
    }

    [TestMethod]
    public void ShouldSetNullToNull()
    {
        _ = this._table.Rows.Add("1", null);

        _ = this._worksheet.Cells["A1"]
            .LoadFromDataTable(this._table,
                               c =>
                               {
                                   c.PrintHeaders = true;
                                   c.TableStyle = TableStyles.Dark1;
                               });

        Assert.IsNull(this._worksheet.Cells["B2"].Value);
    }

    [TestMethod]
    public void ShouldReplaceWithNullIfDbNull()
    {
        _ = this._table.Rows.Add("1", null);
        this._worksheet.Cells["B2"].Value = 2;

        _ = this._worksheet.Cells["A1"]
            .LoadFromDataTable(this._table,
                               c =>
                               {
                                   c.PrintHeaders = true;
                                   c.TableStyle = TableStyles.Dark1;
                               });

        Assert.IsNull(this._worksheet.Cells["B2"].Value);
    }
}