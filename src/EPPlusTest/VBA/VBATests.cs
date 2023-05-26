﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using OfficeOpenXml.VBA.ContentHash;
using OfficeOpenXml.VBA.Signatures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.VBA
{
    [TestClass]
    public class VBATests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Vba.xlsm", true);
            _pck.Workbook.CreateVBAProject();
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void Compression()
        {
            //Compression/Decompression
            string value = "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa";

            byte[] compValue = VBACompression.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            string decompValue = Encoding.GetEncoding(1252).GetString(VBACompression.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);

            value = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

            compValue = VBACompression.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            decompValue = Encoding.GetEncoding(1252).GetString(VBACompression.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);
        }
        [TestMethod]
        public void WriteVBA()
        {
            _pck.Workbook.Worksheets.Add("WriteVBA");
            _pck.Workbook.VbaProject.Modules["WriteVBA"].Code += "\r\nPrivate Sub Worksheet_SelectionChange(ByVal Target As Range)\r\nMsgBox(\"Test of the VBA Feature!\")\r\nEnd Sub\r\n";
            _pck.Workbook.VbaProject.Modules["WriteVBA"].Name = "Blad1";
            _pck.Workbook.CodeModule.Name = "DenHärArbetsboken";
            _pck.Workbook.Worksheets[0].Name = "FirstSheet";
            //_pck.Workbook.CodeModule.Code += "\r\nPrivate Sub Workbook_Open()\r\nBlad1.Cells(1,1).Value = \"VBA test\"\r\nMsgBox \"VBA is running!\"\r\nEnd Sub";
            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[11];

            ExcelVBAModule? m = _pck.Workbook.VbaProject.Modules.AddModule("Module1");
            m.Code += "Public Sub Test(param1 as string)\r\n\r\nEnd sub\r\nPublic Function functest() As String\r\n\r\nEnd Function\r\n";
            ExcelVBAModule? c = _pck.Workbook.VbaProject.Modules.AddClass("Class1", false);
            c.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";
            ExcelVBAModule? c2 = _pck.Workbook.VbaProject.Modules.AddClass("Class2", true);
            c2.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";

            _pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");
        }
        [TestMethod]
        public void WriteLongVBAModule()
        {
            _pck.Workbook.Worksheets.Add("VBASetData");
            _pck.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nCreateData\r\nEnd Sub";
            ExcelVBAModule? module = _pck.Workbook.VbaProject.Modules.AddModule("Code");

            StringBuilder code = new StringBuilder("Public Sub CreateData()\r\n");
            for (int row = 1; row < 30; row++)
            {
                for (int col = 1; col < 30; col++)
                {
                    code.AppendLine(string.Format("VBASetData.Cells({0},{1}).Value=\"Cell {2}\"", row, col, new ExcelAddressBase(row, col, row, col).Address));
                }
            }
            code.AppendLine("End Sub");
            module.Code = code.ToString();

            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
        }
        [TestMethod]
        public void CreateUnicodeWsName()
        {
            ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("测试");

            StringBuilder? sb = new StringBuilder();
            sb.AppendLine("Sub GetData()");
            sb.AppendLine("MsgBox (\"Hello,World\")");
            sb.AppendLine("End Sub");

            ExcelWorksheet worksheet2 = _pck.Workbook.Worksheets.Add("Sheet1");
            StringBuilder? stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Private Sub Worksheet_Change(ByVal Target As Range)");
            stringBuilder.AppendLine("GetData");
            stringBuilder.AppendLine("End Sub");
            worksheet.CodeModule.Code = stringBuilder.ToString();
        }

        [TestMethod]
        public void ValidateName()
        {
            using ExcelPackage? p = new ExcelPackage();
            p.Workbook.CreateVBAProject();
            p.Workbook.Worksheets.Add("Work!Sheet");
            p.Workbook.Worksheets.Add("Mod=ule1");
            p.Workbook.Worksheets.Add("_module1");
            p.Workbook.Worksheets.Add("1module1");
            p.Workbook.Worksheets.Add("Module_1");

            Assert.AreEqual("ThisWorkbook", p.Workbook.VbaProject.Modules[0].Name);
            Assert.AreEqual("Sheet0", p.Workbook.VbaProject.Modules[1].Name);
            Assert.AreEqual("Sheet1", p.Workbook.VbaProject.Modules[2].Name);
            Assert.AreEqual("Sheet2", p.Workbook.VbaProject.Modules[3].Name);
            Assert.AreEqual("Sheet3", p.Workbook.VbaProject.Modules[4].Name);
            Assert.AreEqual("Module_1", p.Workbook.VbaProject.Modules[5].Name);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ModuleNameContainsInvalidCharacters()
        {
            using ExcelPackage? p = new ExcelPackage();
            p.Workbook.Worksheets.Add("InvalidName");
            p.Workbook.CreateVBAProject();
            p.Workbook.VbaProject.Modules.AddModule("Mod%ule2");
        }
        [TestMethod]
        public void ValidateModuleNameAfterCopyWorksheet()
        {
            using ExcelPackage? p = new ExcelPackage();
            string? wsName = "SheetWithLooooooooooooooongName";
            ExcelWorksheet? ws = p.Workbook.Worksheets.Add(wsName);
            p.Workbook.CreateVBAProject();
            ws.CodeModule.Code = "Sub VBA_Code\r\n\r\nEnd Sub";

            ExcelWorksheet? newWS1 = p.Workbook.Worksheets.Add("1newworksheet", ws);
            ExcelWorksheet? newWS2 = p.Workbook.Worksheets.Add("Sheet3", ws);
            ExcelWorksheet? newWS3 = p.Workbook.Worksheets.Add("newworksheet+1", ws);

            Assert.AreEqual(5, p.Workbook.VbaProject.Modules.Count);
            Assert.AreEqual("ThisWorkbook", p.Workbook.VbaProject.Modules[0].Name);
            Assert.AreEqual(wsName, p.Workbook.VbaProject.Modules[1].Name);
            Assert.AreEqual("Sheet1", p.Workbook.VbaProject.Modules[2].Name);
            Assert.AreEqual("Sheet3", p.Workbook.VbaProject.Modules[3].Name);
            Assert.AreEqual("Sheet4", p.Workbook.VbaProject.Modules[4].Name);
        }

        [TestMethod]
        public void SignedUnsignedWorkbook()
        {
            using ExcelPackage? package = OpenTemplatePackage(@"SignedUnsignedWorkbook1.xlsm");
            ExcelVbaProject? proj = package.Workbook.VbaProject;
            ExcelVbaSignature? s = proj.Signature;
            s.LegacySignature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA512;
            s.AgileSignature.CreateSignatureOnSave = false;
            s.V3Signature.CreateSignatureOnSave = false;
            SaveWorkbook("SavedSignedUnsignedWorkbook1.xlsm", package);
        }
        [TestMethod]
        public void Verify_SignedWorkbook1_Hash_V3()
        {
            using ExcelPackage? package = OpenTemplatePackage(@"SignedWorkbook1.xlsm");
            ExcelVbaProject? proj = package.Workbook.VbaProject;
            ExcelVbaSignature? s = proj.Signature;
            EPPlusSignatureContext? ctx = s.V3Signature.SignatureHandler.Context;

            byte[]? hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);
            Assert.IsTrue(ctx.SourceHash.SequenceEqual(hash));
        }
        
        [TestMethod]
        public void Verify_SignedWorkbook1_Hash_Agile()
        {
            using ExcelPackage? package = OpenTemplatePackage(@"SignedWorkbook1.xlsm");
            ExcelVbaProject? proj = package.Workbook.VbaProject;
            ExcelVbaSignature? s = proj.Signature;
            EPPlusSignatureContext? ctx = s.AgileSignature.SignatureHandler.Context;

            byte[]? hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);
            Assert.IsTrue(ctx.SourceHash.SequenceEqual(hash));
        }
        [TestMethod]
        public void Verify_SignedWorkbook1_Hash_Legacy()
        {
            using ExcelPackage? package = OpenTemplatePackage(@"SignedWorkbook1.xlsm");
            ExcelVbaProject? proj = package.Workbook.VbaProject;
            ExcelVbaSignature? s = proj.Signature;
            EPPlusSignatureContext? ctx = s.LegacySignature.SignatureHandler.Context;

            byte[]? hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);
            Assert.IsTrue(ctx.SourceHash.SequenceEqual(hash));
        }
        [TestMethod]
        public void SignedWorkbook()
        {
            using ExcelPackage? package = OpenTemplatePackage(@"SignedWorkbook1.xlsm");
            ExcelVbaProject? proj = package.Workbook.VbaProject;
            ExcelVbaSignature? s = proj.Signature;
            package.Workbook.VbaProject.Signature.LegacySignature.CreateSignatureOnSave = false;
            package.Workbook.VbaProject.Signature.V3Signature.CreateSignatureOnSave = false;
            SaveAndCleanup(package);
        }
        [TestMethod]
        public void MyVbaTest_Sign1()
        {
            string? workbook = "VbaSignedSimple2.xlsm";
            using ExcelPackage? package = OpenTemplatePackage(workbook);
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            foreach (X509Certificate2? cert in store.Certificates)
            {
                if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                {
                    if (cert.Thumbprint == "C0201D22A64D78757EF4655988B267E6734E04B5")
                    {
                        package.Workbook.VbaProject.Signature.Certificate = cert;
                        break;
                    }
                }
            }
            ExcelVBAModule? module = package.Workbook.VbaProject.Modules.AddModule("TestCode");
            module.Code = "Sub Main\r\nMsgbox(\"Test\")\r\nEnd Sub";
            package.Workbook.VbaProject.Signature.LegacySignature.CreateSignatureOnSave = false;
            package.Workbook.VbaProject.Signature.V3Signature.CreateSignatureOnSave = false;
            package.Workbook.VbaProject.Signature.AgileSignature.HashAlgorithm = OfficeOpenXml.VBA.VbaSignatureHashAlgorithm.SHA256;
            SaveWorkbook("SignedUnsignedWorkbook1.xlsm", package);
        }
        [TestMethod]
        public void v3ContentSigningSample()
        {
            string? workbook = "v3Signing\\V3ContentSigning_original.xlsm";
            using ExcelPackage? package = OpenTemplatePackage(workbook);
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            foreach (X509Certificate2? cert in store.Certificates)
            {
                if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                {
                    if (cert.Thumbprint == "C0201D22A64D78757EF4655988B267E6734E04B5")
                    {
                        package.Workbook.VbaProject.Signature.Certificate = cert;
                        break;
                    }
                }
            }
            package.Workbook.VbaProject.Signature.LegacySignature.CreateSignatureOnSave = false;
            package.Workbook.VbaProject.Signature.AgileSignature.CreateSignatureOnSave = false;
            SaveWorkbook("v3Signing\\EPPlusV3ContentSigning.xlsm", package);
        }
    }
}
