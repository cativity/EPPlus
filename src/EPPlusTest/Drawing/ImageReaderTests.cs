using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace EPPlusTest.Drawing;

[TestClass]
public class ImageReaderTests : TestBase
{
    private static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        InitBase();
        _pck = OpenPackage("ImageReader.xlsx", true);
        _pck.Settings.ImageSettings.PrimaryImageHandler = new GenericImageHandler();
    }

    [ClassCleanup]
    public static void Cleanup()
    {
        string? dirName = _pck.File.DirectoryName;
        string? fileName = _pck.File.FullName;

        SaveAndCleanup(_pck);

        if (File.Exists(fileName))
        {
            File.Copy(fileName, dirName + "\\ImageReaderRead.xlsx", true);
        }
    }

    [TestMethod]
    public void AddJpgImageVia()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("InternalJpg");

        using MemoryStream? ms = new MemoryStream(Properties.Resources.Test1JpgByteArray);
        _ = ws.Drawings.AddPicture("jpg", ms, ePictureType.Jpg);
    }

    [TestMethod]
    public void AddPngImageVia()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("InternalPng");

        using MemoryStream? ms = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray);
        _ = ws.Drawings.AddPicture("png1", ms, ePictureType.Png);
    }

    [TestMethod]
    public void AddTestImagesToWorksheet()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("picturesIS");

        using (MemoryStream? msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
        {
            ExcelPicture? imageGif = ws.Drawings.AddPicture("gif1", msGif, ePictureType.Gif);
            imageGif.SetPosition(40, 0, 0, 0);
        }

        using (MemoryStream? msBmp = new MemoryStream(Properties.Resources.CodeBmp))
        {
            ExcelPicture? imagebmp = ws.Drawings.AddPicture("bmp1", msBmp, ePictureType.Bmp);
            imagebmp.SetPosition(40, 0, 10, 0);
        }

        using (MemoryStream? ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
        {
            _ = ws.Drawings.AddPicture("jpg1", ms1, ePictureType.Jpg);
        }

        using (MemoryStream? ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
        {
            ExcelPicture? image2 = ws.Drawings.AddPicture("png1", ms2, ePictureType.Png);
            image2.SetPosition(0, 0, 10, 0);
        }

        using (MemoryStream? ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
        {
            ExcelPicture? image22 = ws.Drawings.AddPicture("png2", ms22, ePictureType.Png);
            image22.SetPosition(0, 0, 20, 0);
        }

        using (MemoryStream? ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
        {
            ExcelPicture? image23 = ws.Drawings.AddPicture("png3", ms23, ePictureType.Png);
            image23.SetPosition(0, 0, 30, 0);
        }

        using (MemoryStream? ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
        {
            ExcelPicture? image3 = ws.Drawings.AddPicture("emf1", ms3, ePictureType.Emf);
            image3.SetPosition(0, 0, 40, 0);
        }

        using (MemoryStream? ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
        {
            ExcelPicture? image4 = ws.Drawings.AddPicture("svg1", ms4, ePictureType.Svg);
            image4.SetPosition(0, 0, 50, 0);
        }

        using (MemoryStream? ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
        {
            ExcelPicture? image5 = ws.Drawings.AddPicture("svg2", ms5, ePictureType.Svg);
            image5.SetPosition(0, 0, 60, 0);
            image5.SetSize(25);
        }

        using (MemoryStream? ms6 = Properties.Resources.VectorDrawing)
        {
            ExcelPicture? image6 = ws.Drawings.AddPicture("wmf", ms6, ePictureType.Wmf);
            image6.SetPosition(0, 0, 70, 0);
        }

        using MemoryStream? msTif = Properties.Resources.CodeTif;
        ExcelPicture? imageTif = ws.Drawings.AddPicture("tif1", msTif, ePictureType.Tif);
        imageTif.SetPosition(0, 0, 80, 0);
    }

    [TestMethod]
    public void AddTestImagesToWorksheetNoPictureType()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("picturesISNoPT");

        using (MemoryStream? msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
        {
            ExcelPicture? imageGif = ws.Drawings.AddPicture("gif1", msGif);
            Assert.AreEqual("image/gif", imageGif.ContentType);
            imageGif.SetPosition(40, 0, 0, 0);
        }

        using (MemoryStream? msBmp = new MemoryStream(Properties.Resources.CodeBmp))
        {
            ExcelPicture? imagebmp = ws.Drawings.AddPicture("bmp1", msBmp);
            Assert.AreEqual("image/bmp", imagebmp.ContentType);
            imagebmp.SetPosition(40, 0, 10, 0);
        }

        using (MemoryStream? ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
        {
            ExcelPicture? image1 = ws.Drawings.AddPicture("jpg1", ms1);
            Assert.AreEqual("image/jpeg", image1.ContentType);
        }

        using (MemoryStream? ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
        {
            ExcelPicture? image2 = ws.Drawings.AddPicture("png1", ms2);
            image2.SetPosition(0, 0, 10, 0);
            Assert.AreEqual("image/png", image2.ContentType);
        }

        using (MemoryStream? ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
        {
            ExcelPicture? image22 = ws.Drawings.AddPicture("png2", ms22);
            image22.SetPosition(0, 0, 20, 0);
            Assert.AreEqual("image/png", image22.ContentType);
        }

        using (MemoryStream? ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
        {
            ExcelPicture? image23 = ws.Drawings.AddPicture("png3", ms23);
            image23.SetPosition(0, 0, 30, 0);
            Assert.AreEqual("image/png", image23.ContentType);
        }

        using (MemoryStream? ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
        {
            ExcelPicture? image3 = ws.Drawings.AddPicture("emf1", ms3);
            image3.SetPosition(0, 0, 40, 0);
            Assert.AreEqual("image/x-emf", image3.ContentType);
        }

        using (MemoryStream? ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
        {
            ExcelPicture? image4 = ws.Drawings.AddPicture("svg1", ms4);
            image4.SetPosition(0, 0, 50, 0);
            Assert.AreEqual("image/svg+xml", image4.ContentType);
        }

        using (MemoryStream? ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
        {
            ExcelPicture? image5 = ws.Drawings.AddPicture("svg2", ms5);
            image5.SetPosition(0, 0, 60, 0);
            image5.SetSize(25);
            Assert.AreEqual("image/svg+xml", image5.ContentType);
        }

        using (MemoryStream? ms6 = Properties.Resources.VectorDrawing)
        {
            ExcelPicture? image6 = ws.Drawings.AddPicture("wmf", ms6);
            image6.SetPosition(0, 0, 70, 0);
            Assert.AreEqual("image/x-wmf", image6.ContentType);
        }

        using (MemoryStream? msTif = Properties.Resources.CodeTif)
        {
            ExcelPicture? imageTif = ws.Drawings.AddPicture("tif1", msTif);
            imageTif.SetPosition(0, 0, 80, 0);
            Assert.AreEqual("image/x-tiff", imageTif.ContentType);
        }

        using MemoryStream? msIco128 = GetImageMemoryStream("1_128x128.ico");
        ExcelPicture? imageIco = ws.Drawings.AddPicture("ico2", msIco128, ePictureType.Ico);
        imageIco.SetPosition(40, 0, 10, 0);
        Assert.AreEqual("image/x-icon", imageIco.ContentType);
    }

    [TestMethod]
    public async Task AddTestImagesToWorksheetNoPictureTypeAsync()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("picturesISNoPTAsync");

        using (MemoryStream? msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
        {
            ExcelPicture? imageGif = await ws.Drawings.AddPictureAsync("gif1", msGif);
            Assert.AreEqual("image/gif", imageGif.ContentType);
            imageGif.SetPosition(40, 0, 0, 0);
        }

        using (MemoryStream? msBmp = new MemoryStream(Properties.Resources.CodeBmp))
        {
            ExcelPicture? imagebmp = await ws.Drawings.AddPictureAsync("bmp1", msBmp);
            Assert.AreEqual("image/bmp", imagebmp.ContentType);
            imagebmp.SetPosition(40, 0, 10, 0);
        }

        using (MemoryStream? ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
        {
            ExcelPicture? image1 = await ws.Drawings.AddPictureAsync("jpg1", ms1);
            Assert.AreEqual("image/jpeg", image1.ContentType);
        }

        using (MemoryStream? ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
        {
            ExcelPicture? image2 = await ws.Drawings.AddPictureAsync("png1", ms2);
            image2.SetPosition(0, 0, 10, 0);
            Assert.AreEqual("image/png", image2.ContentType);
        }

        using (MemoryStream? ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
        {
            ExcelPicture? image22 = await ws.Drawings.AddPictureAsync("png2", ms22);
            image22.SetPosition(0, 0, 20, 0);
            Assert.AreEqual("image/png", image22.ContentType);
        }

        using (MemoryStream? ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
        {
            ExcelPicture? image23 = await ws.Drawings.AddPictureAsync("png3", ms23);
            image23.SetPosition(0, 0, 30, 0);
            Assert.AreEqual("image/png", image23.ContentType);
        }

        using (MemoryStream? ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
        {
            ExcelPicture? image3 = await ws.Drawings.AddPictureAsync("emf1", ms3);
            image3.SetPosition(0, 0, 40, 0);
            Assert.AreEqual("image/x-emf", image3.ContentType);
        }

        using (MemoryStream? ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
        {
            ExcelPicture? image4 = await ws.Drawings.AddPictureAsync("svg1", ms4);
            image4.SetPosition(0, 0, 50, 0);
            Assert.AreEqual("image/svg+xml", image4.ContentType);
        }

        using (MemoryStream? ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
        {
            ExcelPicture? image5 = await ws.Drawings.AddPictureAsync("svg2", ms5);
            image5.SetPosition(0, 0, 60, 0);
            image5.SetSize(25);
            Assert.AreEqual("image/svg+xml", image5.ContentType);
        }

        using (MemoryStream? ms6 = Properties.Resources.VectorDrawing)
        {
            ExcelPicture? image6 = await ws.Drawings.AddPictureAsync("wmf", ms6);
            image6.SetPosition(0, 0, 70, 0);
            Assert.AreEqual("image/x-wmf", image6.ContentType);
        }

        using (MemoryStream? msTif = Properties.Resources.CodeTif)
        {
            ExcelPicture? imageTif = await ws.Drawings.AddPictureAsync("tif1", msTif);
            imageTif.SetPosition(0, 0, 80, 0);
            Assert.AreEqual("image/x-tiff", imageTif.ContentType);
        }

        using MemoryStream? msIco128 = GetImageMemoryStream("1_128x128.ico");
        ExcelPicture? imageIco = await ws.Drawings.AddPictureAsync("ico2", msIco128, ePictureType.Ico);
        imageIco.SetPosition(40, 0, 10, 0);
        Assert.AreEqual("image/x-icon", imageIco.ContentType);
    }

    [TestMethod]
    public void AddIcoImages()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Icon");

        //32*32
        using (MemoryStream? msIco1 = GetImageMemoryStream("1_32x32.ico"))
        {
            ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("ico1", msIco1, ePictureType.Ico);
            imageWebP1.SetPosition(40, 0, 0, 0);
        }

        //128*128
        using (MemoryStream? msIco2 = GetImageMemoryStream("1_128x128.ico"))
        {
            ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("ico2", msIco2, ePictureType.Ico);
            imageWebP1.SetPosition(40, 0, 10, 0);
        }

        using (MemoryStream? msIco3 = GetImageMemoryStream("example.ico"))
        {
            ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("ico3", msIco3, ePictureType.Ico);
            imageWebP1.SetPosition(40, 0, 20, 0);
        }

        using (MemoryStream? msIco4 = GetImageMemoryStream("example_small.ico"))
        {
            ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("ico4", msIco4, ePictureType.Ico);
            imageWebP1.SetPosition(40, 0, 30, 0);
        }

        using (MemoryStream? msIco5 = GetImageMemoryStream("Ico-file-for-testing.ico"))
        {
            ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("ico5", msIco5, ePictureType.Ico);
            imageWebP1.SetPosition(40, 0, 40, 0);
        }
    }

    [TestMethod]
    public void AddEmzImages()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("Emz");

        //32*32
        using MemoryStream? msIco1 = GetImageMemoryStream("example.emz");
        ExcelPicture? imageWebP1 = ws.Drawings.AddPicture("Emf", msIco1, ePictureType.Emz);
        imageWebP1.SetPosition(40, 0, 0, 0);
    }

    [TestMethod]
    public void AddBmpImages()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("bmp");

        using (MemoryStream? msBmp1 = GetImageMemoryStream("bmp\\MARBLES.BMP"))
        {
            ExcelPicture? imageBmp1 = ws.Drawings.AddPicture("bmp1", msBmp1, ePictureType.Bmp);
            imageBmp1.SetPosition(0, 0, 0, 0);
        }

        using (MemoryStream? msBmp2 = GetImageMemoryStream("bmp\\Land.BMP"))
        {
            ExcelPicture? imageBmp2 = ws.Drawings.AddPicture("bmp2", msBmp2, ePictureType.Bmp);
            imageBmp2.SetPosition(0, 0, 20, 0);
        }

        using (MemoryStream? msBmp3 = GetImageMemoryStream("bmp\\Land2.BMP"))
        {
            ExcelPicture? imageBmp3 = ws.Drawings.AddPicture("bmp3", msBmp3, ePictureType.Bmp);
            imageBmp3.SetPosition(0, 0, 40, 0);
        }

        using MemoryStream? msBmp4 = GetImageMemoryStream("bmp\\Land3.BMP");
        ExcelPicture? imageBmp4 = ws.Drawings.AddPicture("bmp4", msBmp4, ePictureType.Bmp);
        imageBmp4.SetPosition(0, 0, 60, 0);
    }

    [TestMethod]
    public void AddWepPImages() => AddFilesToWorksheet("webp", ePictureType.WebP);

    [TestMethod]
    public void AddWepPImagesNoPT() => AddFilesToWorksheet("webp", null, "webp-NoPT");

    [TestMethod]
    public void AddEmfImages() => AddFilesToWorksheet("Emf", ePictureType.Emf);

    [TestMethod]
    public void AddGifImages() => AddFilesToWorksheet("Gif", ePictureType.Gif);

    [TestMethod]
    public void AddJpgImages() => AddFilesToWorksheet("Jpg", ePictureType.Jpg);

    [TestMethod]
    public void AddSvgImages() => AddFilesToWorksheet("Svg", ePictureType.Svg);

    [TestMethod]
    public void AddPngImages() => AddFilesToWorksheet("Png", ePictureType.Png);

    [TestMethod]
    public void ReadImages()
    {
        using ExcelPackage? p = OpenPackage("ImageReaderRead.xlsx");

        if (p.Workbook.Worksheets.Count == 0)
        {
            Assert.Inconclusive("ImageReaderRead.xlsx does not exists. Run a full test round to create it.");
        }

        foreach (ExcelWorksheet? ws in p.Workbook.Worksheets)
        {
            ws.Columns[1, 20].Width = 35;

            Assert.AreEqual(35, ws.Columns[1].Width);
            Assert.AreEqual(35, ws.Columns[20].Width);
        }

        ExcelWorksheet? ws2 = p.Workbook.Worksheets.Add("Bmp2");

        using (MemoryStream? msBmp1 = GetImageMemoryStream("bmp\\MARBLES.BMP"))
        {
            ExcelPicture? imageBmp1 = ws2.Drawings.AddPicture("bmp2", msBmp1, ePictureType.Bmp);
            imageBmp1.SetPosition(0, 0, 0, 0);
        }

        SaveWorkbook("ImageReaderResized.xlsx", p);
    }

    [TestMethod]
    public async Task AddJpgImagesViaExcelImage()
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add("AddViaExcelImage");

        ExcelImage? ei1 = new ExcelImage(Properties.Resources.Test1.FullName);
        Assert.IsNotNull(ei1);
        _ = ws.BackgroundImage.Image.SetImage(ei1);

        ExcelImage? ei2 = new ExcelImage(Properties.Resources.Png2ByteArray, ePictureType.Png);
        Assert.IsNotNull(ei2);
        _ = ws.BackgroundImage.Image.SetImage(ei2);

        ExcelImage? ei3 = new ExcelImage(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
        Assert.IsNotNull(ei3);

        _ = ws.BackgroundImage.Image.SetImage(ei3);
        _ = ws.BackgroundImage.Image.SetImage(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
        _ = await ws.BackgroundImage.Image.SetImageAsync(new MemoryStream(Properties.Resources.BitmapImageGif), ePictureType.Gif);
    }

    private static void AddFilesToWorksheet(string fileType, ePictureType? type, string worksheetName = null)
    {
        ExcelWorksheet? ws = _pck.Workbook.Worksheets.Add(worksheetName ?? fileType);

        DirectoryInfo? dir = new DirectoryInfo(_imagePath + fileType);

        if (dir.Exists == false)
        {
            Assert.Inconclusive($"Directory {dir} does not exist.");
        }

        int ix = 0;

        foreach (FileInfo? f in dir.EnumerateFiles())
        {
            using MemoryStream? ms = new MemoryStream(File.ReadAllBytes(f.FullName));
            ExcelPicture? picture = ws.Drawings.AddPicture($"{fileType}{ix}", ms, type);
            picture.SetPosition(ix / 5 * 10, 0, ix % 5 * 10, 0);
            ix++;
        }
    }
}