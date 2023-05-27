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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using System.Drawing;

namespace EPPlusTest.Drawing;

[TestClass]
public class FillReadTest : TestBase
{
    static ExcelPackage _pck;

    [ClassInitialize]
    public static void Init(TestContext context)
    {
        _pck = OpenPackage("DrawingFillRead.xlsx");
    }

    #region SolidFill

    [TestMethod]
    public void ReadColorProperty()
    {
        //Setup
        string? wsName = "SolidFill";
        Color expected = Color.Blue;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
        Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
    }

    [TestMethod]
    public void ReadSolidFill_Color()
    {
        //Setup
        string? wsName = "SolidFillFromSolidFill";
        Color expected = Color.Green;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbColor);
        Assert.AreEqual(expected.ToArgb(), shape.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
    }

    [TestMethod]
    public void ReadSolidFill_ColorPreset()
    {
        //Setup
        string? wsName = "SolidFillFromPresetClr";
        ePresetColor expected = ePresetColor.Red;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.PresetColor);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.PresetColor.Color);
    }

    [TestMethod]
    public void ReadSolidFill_ColorScheme()
    {
        //Setup
        string? wsName = "SolidFillFromSchemeClr";
        eSchemeColor expected = eSchemeColor.Accent6;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Act
        shape.Fill.Style = eFillStyle.SolidFill;
        shape.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.SchemeColor);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SchemeColor.Color);
    }

    [TestMethod]
    public void ReadSolidFill_ColorPercentage()
    {
        //Setup
        string? wsName = "SolidFillFromColorPrc";
        int expectedR = 51;
        int expectedG = 49;
        int expectedB = 50;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.RgbPercentageColor);
        Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.RgbPercentageColor.RedPercentage);
        Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.RgbPercentageColor.GreenPercentage);
        Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.RgbPercentageColor.BluePercentage);
    }

    [TestMethod]
    public void ReadSolidFill_ColorHsl()
    {
        //Setup
        string? wsName = "SolidFillFromColorHcl";
        int expectedHue = 180;
        int expectedLum = 15;
        int expectedSat = 50;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.HslColor);
        Assert.AreEqual(expectedHue, shape.Fill.SolidFill.Color.HslColor.Hue);
        Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.HslColor.Luminance);
        Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.HslColor.Saturation);
    }

    [TestMethod]
    public void ReadSolidFill_ColorSystem()
    {
        //Setup
        string? wsName = "SolidFillFromColorSystem";
        eSystemColor expected = eSystemColor.Background;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.IsNotNull(shape.Fill.SolidFill.Color.SystemColor);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.SystemColor.Color);
    }

    #endregion

    #region Transform

    [TestMethod]
    public void ReadTransparancy()
    {
        //Setup
        string? wsName = "Transparancy";
        int expected = 45;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(expected, shape.Fill.Transparancy);
        Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(100 - expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
    }

    [TestMethod]
    public void ReadTransformAlpha()
    {
        //Setup
        string? wsName = "Alpha";
        int expected = 45;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(100 - expected, shape.Fill.Transparancy);
        Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
    }

    [TestMethod]
    public void ReadTransformTint()
    {
        //Setup
        string? wsName = "Tint";
        int expected = 30;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Tint, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
    }

    [TestMethod]
    public void ReadTransformShade()
    {
        //Setup
        string? wsName = "Shade";
        int expected = 95;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Shade, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[0].Value);
    }

    [TestMethod]
    public void ReadTransformInverse_true()
    {
        //Setup
        string? wsName = "Inverse_set";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Inv, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(1, shape.Fill.SolidFill.Color.Transforms[0].Value);
    }

    [TestMethod]
    public void ReadTransformAlphaModulation()
    {
        //Setup
        string? wsName = "AlphaModulation";
        int expected = 50;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(eColorTransformType.AlphaMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(20, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
    }

    [TestMethod]
    public void ReadTransformAlphaOffset()
    {
        //Setup
        string? wsName = "AlphaOffset";
        int expected = -10;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Alpha, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(eColorTransformType.AlphaOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(20, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(expected, shape.Fill.SolidFill.Color.Transforms[1].Value);
    }

    [TestMethod]
    public void ReadTransformColorPercentage()
    {
        //Setup
        string? wsName = "TransColorPerc";
        int expectedR = 30;
        int expectedG = 60;
        int expectedB = 20;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.Red, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(eColorTransformType.Green, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
        Assert.AreEqual(eColorTransformType.Blue, shape.Fill.SolidFill.Color.Transforms[2].Type);
        Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
    }

    [TestMethod]
    public void ReadTransformColorModulation()
    {
        //Setup
        string? wsName = "TransColorMod";
        double expectedR = 3.33;
        int expectedG = 50;
        int expectedB = 25600;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.RedMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(eColorTransformType.GreenMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
        Assert.AreEqual(eColorTransformType.BlueMod, shape.Fill.SolidFill.Color.Transforms[2].Type);
        Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
    }

    [TestMethod]
    public void ReadTransformColorOffset()
    {
        //Setup
        string? wsName = "TransColorOffset";
        int expectedR = 10;
        int expectedG = -20;
        int expectedB = 30;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.RedOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expectedR, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(eColorTransformType.GreenOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(expectedG, shape.Fill.SolidFill.Color.Transforms[1].Value);
        Assert.AreEqual(eColorTransformType.BlueOff, shape.Fill.SolidFill.Color.Transforms[2].Type);
        Assert.AreEqual(expectedB, shape.Fill.SolidFill.Color.Transforms[2].Value);
    }

    [TestMethod]
    public void ReadTransformHslOffset()
    {
        //Setup
        string? wsName = "TransHslOffset";
        int expectedLum = 10;
        int expectedSat = -20;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.LumOff, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(eColorTransformType.SatOff, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
    }

    [TestMethod]
    public void ReadTransformHslModulation()
    {
        //Setup
        string? wsName = "TransHslModulation";
        int expectedLum = 50;
        int expectedSat = 200;
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.SolidFill, shape.Fill.Style);
        Assert.AreEqual(eColorTransformType.LumMod, shape.Fill.SolidFill.Color.Transforms[0].Type);
        Assert.AreEqual(expectedLum, shape.Fill.SolidFill.Color.Transforms[0].Value);
        Assert.AreEqual(eColorTransformType.SatMod, shape.Fill.SolidFill.Color.Transforms[1].Type);
        Assert.AreEqual(expectedSat, shape.Fill.SolidFill.Color.Transforms[1].Value);
    }

    #endregion

    #region Gradient

    [TestMethod]
    public void ReadGradient()
    {
        //Setup
        string? wsName = "Gradient";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
        Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);
        Assert.AreEqual(true, shape.Fill.GradientFill.RotateWithShape);
        Assert.AreEqual(eTileFlipMode.XY, shape.Fill.GradientFill.TileFlip);
    }

    [TestMethod]
    public void ReadGradientCircularPath()
    {
        //Setup
        string? wsName = "GradientCircular";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);
        Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
        Assert.AreEqual(eShadePath.Circle, shape.Fill.GradientFill.ShadePath);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
    }

    [TestMethod]
    public void ReadGradientRectPath()
    {
        //Setup
        string? wsName = "GradientRect";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.AreEqual(eShadePath.Rectangle, shape.Fill.GradientFill.ShadePath);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
        Assert.AreEqual(20, shape.Fill.GradientFill.FocusPoint.BottomOffset);
        Assert.AreEqual(20, shape.Fill.GradientFill.FocusPoint.LeftOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
    }

    [TestMethod]
    public void ReadGradientShapePath()
    {
        //Setup
        string? wsName = "GradientShape";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.AreEqual(3, shape.Fill.GradientFill.Colors.Count);

        Assert.AreEqual(Color.LightBlue.ToArgb(), shape.Fill.GradientFill.Colors[0D].Color.RgbColor.Color.ToArgb());
        Assert.AreEqual(Color.Blue.ToArgb(), shape.Fill.GradientFill.Colors[40D].Color.RgbColor.Color.ToArgb());
        Assert.AreEqual(Color.DarkBlue.ToArgb(), shape.Fill.GradientFill.Colors[100D].Color.RgbColor.Color.ToArgb());
        Assert.IsNull(shape.Fill.GradientFill.Colors[41D]);

        Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
        Assert.AreEqual(eShadePath.Shape, shape.Fill.GradientFill.ShadePath);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.TopOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.BottomOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.LeftOffset);
        Assert.AreEqual(50, shape.Fill.GradientFill.FocusPoint.RightOffset);
    }

    [TestMethod]
    public void ReadGradientAddMethods()
    {
        //Setup
        string? wsName = "GradientAddMethods";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Act
        shape.Fill.Style = eFillStyle.GradientFill;
        shape.Fill.GradientFill.Colors.AddRgb(0, Color.Red);
        shape.Fill.GradientFill.Colors.AddRgbPercentage(22.55, 40, 50, 60.5);
        shape.Fill.GradientFill.Colors.AddHsl(37.42, 180, 50, 60);
        shape.Fill.GradientFill.Colors.AddPreset(55.2, ePresetColor.BlueViolet);
        shape.Fill.GradientFill.Colors.AddScheme(66.2, eSchemeColor.Background2);
        shape.Fill.GradientFill.Colors.AddSystem(88.2, eSystemColor.GradientActiveCaption);

        //Assert
        Assert.AreEqual(eFillStyle.GradientFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.AreEqual(eDrawingColorType.RgbPercentage, shape.Fill.GradientFill.Colors[22.55].Color.ColorType); //Verify index for position

        //RGB
        Assert.AreEqual(0, shape.Fill.GradientFill.Colors[0].Position);
        Assert.AreEqual(eDrawingColorType.Rgb, shape.Fill.GradientFill.Colors[0].Color.ColorType);
        Assert.AreEqual(Color.Red.ToArgb(), shape.Fill.GradientFill.Colors[0].Color.RgbColor.Color.ToArgb());

        //RGB Percent
        Assert.AreEqual(22.55, shape.Fill.GradientFill.Colors[1].Position);
        Assert.AreEqual(eDrawingColorType.RgbPercentage, shape.Fill.GradientFill.Colors[1].Color.ColorType);
        Assert.AreEqual(40, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.RedPercentage);
        Assert.AreEqual(50, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.GreenPercentage);
        Assert.AreEqual(60.5, shape.Fill.GradientFill.Colors[1].Color.RgbPercentageColor.BluePercentage);

        //Hsl Percent
        Assert.AreEqual(37.42, shape.Fill.GradientFill.Colors[2].Position);
        Assert.AreEqual(eDrawingColorType.Hsl, shape.Fill.GradientFill.Colors[2].Color.ColorType);
        Assert.AreEqual(180, shape.Fill.GradientFill.Colors[2].Color.HslColor.Hue);
        Assert.AreEqual(50, shape.Fill.GradientFill.Colors[2].Color.HslColor.Saturation);
        Assert.AreEqual(60, shape.Fill.GradientFill.Colors[2].Color.HslColor.Luminance);

        //Preset
        Assert.AreEqual(55.2, shape.Fill.GradientFill.Colors[3].Position);
        Assert.AreEqual(eDrawingColorType.Preset, shape.Fill.GradientFill.Colors[3].Color.ColorType);
        Assert.AreEqual(ePresetColor.BlueViolet, shape.Fill.GradientFill.Colors[3].Color.PresetColor.Color);

        //Scheme color
        Assert.AreEqual(66.2, shape.Fill.GradientFill.Colors[4].Position);
        Assert.AreEqual(eDrawingColorType.Scheme, shape.Fill.GradientFill.Colors[4].Color.ColorType);
        Assert.AreEqual(eSchemeColor.Background2, shape.Fill.GradientFill.Colors[4].Color.SchemeColor.Color);

        //Scheme color
        Assert.AreEqual(88.2, shape.Fill.GradientFill.Colors[5].Position);
        Assert.AreEqual(eDrawingColorType.System, shape.Fill.GradientFill.Colors[5].Color.ColorType);
        Assert.AreEqual(eSystemColor.GradientActiveCaption, shape.Fill.GradientFill.Colors[5].Color.SystemColor.Color);
    }

    #endregion

    #region Pattern

    [TestMethod]
    public void ReadPatternDefault()
    {
        //Setup
        string? wsName = "PatternDefault";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.AreEqual(eFillPatternStyle.Pct5, shape.Fill.PatternFill.PatternType);
    }

    [TestMethod]
    public void ReadPatternCross()
    {
        //Setup
        string? wsName = "PatternCross";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Act
        shape.Fill.Style = eFillStyle.PatternFill;
        shape.Fill.PatternFill.PatternType = eFillPatternStyle.Cross;
        shape.Fill.PatternFill.BackgroundColor.SetSchemeColor(eSchemeColor.Accent4);
        shape.Fill.PatternFill.ForegroundColor.SetSchemeColor(eSchemeColor.Background2);

        //Assert
        Assert.AreEqual(eFillStyle.PatternFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.AreEqual(eFillPatternStyle.Cross, shape.Fill.PatternFill.PatternType);
    }

    #endregion

    #region Blip

    [TestMethod]
    public void ReadBlipFill_DefaultSettings()
    {
        //Setup
        string? wsName = "BlipFill";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.IsNull(shape.Fill.PatternFill);
        Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
    }

    [TestMethod]
    public void ReadBlipFill_NoImage()
    {
        //Setup
        string? wsName = "BlipFillNoImage";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.IsNull(shape.Fill.PatternFill);
    }

    [TestMethod]
    public void ReadBlipFill_Stretch()
    {
        //Setup
        string? wsName = "BlipFillStretch";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.IsNull(shape.Fill.PatternFill);
        Assert.AreEqual(true, shape.Fill.BlipFill.Stretch);
        Assert.AreEqual(20, shape.Fill.BlipFill.StretchOffset.TopOffset);
        Assert.AreEqual(10, shape.Fill.BlipFill.StretchOffset.BottomOffset);
        Assert.AreEqual(-5, shape.Fill.BlipFill.StretchOffset.LeftOffset);
        Assert.AreEqual(15, shape.Fill.BlipFill.StretchOffset.RightOffset);
    }

    [TestMethod]
    public void ReadBlipFill_SourceRectangle()
    {
        //Setup
        string? wsName = "BlipFillSourceRectangle";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.IsNull(shape.Fill.PatternFill);
        Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
        Assert.AreEqual(20, shape.Fill.BlipFill.SourceRectangle.TopOffset);
        Assert.AreEqual(10, shape.Fill.BlipFill.SourceRectangle.BottomOffset);
        Assert.AreEqual(-5, shape.Fill.BlipFill.SourceRectangle.LeftOffset);
        Assert.AreEqual(15, shape.Fill.BlipFill.SourceRectangle.RightOffset);
    }

    [TestMethod]
    public void ReadBlipFill_Tile()
    {
        //Setup
        string? wsName = "BlipFillTile";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelShape? shape = (ExcelShape)ws.Drawings[0];

        //Assert
        Assert.AreEqual(eFillStyle.BlipFill, shape.Fill.Style);
        Assert.IsNull(shape.Fill.SolidFill);
        Assert.IsNull(shape.Fill.GradientFill);
        Assert.IsNull(shape.Fill.PatternFill);
        Assert.AreEqual(false, shape.Fill.BlipFill.Stretch);
        Assert.AreEqual(eRectangleAlignment.Center, shape.Fill.BlipFill.Tile.Alignment);
        Assert.AreEqual(eTileFlipMode.XY, shape.Fill.BlipFill.Tile.FlipMode);
        Assert.AreEqual(95, shape.Fill.BlipFill.Tile.HorizontalRatio);
        Assert.AreEqual(97, shape.Fill.BlipFill.Tile.VerticalRatio);
        Assert.AreEqual(2, shape.Fill.BlipFill.Tile.HorizontalOffset);
        Assert.AreEqual(1, shape.Fill.BlipFill.Tile.VerticalOffset);
    }

    [TestMethod]
    public void ReadBlipFill_PieChart()
    {
        //Setup
        string? wsName = "BlipFillPieChart";
        ExcelWorksheet? ws = _pck.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            Assert.Inconclusive($"{wsName} worksheet is missing");
        }

        ExcelPieChart? chart = ws.Drawings[0].As.Chart.PieChart;

        Assert.AreEqual(eFillStyle.BlipFill, chart.Fill.Style);
        Assert.IsNull(chart.Fill.SolidFill);
        Assert.IsNull(chart.Fill.GradientFill);
        Assert.IsNull(chart.Fill.PatternFill);
        Assert.IsNotNull(chart.Fill.BlipFill.Image);
    }

    #endregion
}