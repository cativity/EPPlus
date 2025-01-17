﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/

using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Image;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.SystemDrawing.Image;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml;

/// <summary>
/// This class contains settings for text measurement.
/// </summary>
public class ExcelImageSettings
{
    internal ExcelImageSettings()
    {
        SystemDrawingImageHandler? m = new SystemDrawingImageHandler();

        if (m.ValidForEnvironment())
        {
            this.PrimaryImageHandler = m;
            this.SecondaryImageHandler = new GenericImageHandler();
            this.TertiaryImageHandler = null;
        }
        else
        {
            this.PrimaryImageHandler = new GenericImageHandler();
            this.SecondaryImageHandler = null;
            this.TertiaryImageHandler = null;
        }
    }

    /// <summary>
    /// This is the primary handler for images.
    /// </summary>
    public IImageHandler PrimaryImageHandler { get; set; }

    /// <summary>
    /// If the primary handler fails to measure the image, this one will be used.
    /// </summary>
    public IImageHandler SecondaryImageHandler { get; set; }

    /// <summary>
    /// If the secondary handler fails to measure the image, this one will be used.
    /// </summary>
    public IImageHandler TertiaryImageHandler { get; set; }

    internal bool GetImageBounds(MemoryStream ms,
                                 ePictureType type,
                                 out double width,
                                 out double height,
                                 out double horizontalResolution,
                                 out double verticalResolution)
    {
        if (this.PrimaryImageHandler.SupportedTypes.Contains(type)
            && this.PrimaryImageHandler.GetImageBounds(ms, type, out width, out height, out horizontalResolution, out verticalResolution))
        {
            return true;
        }

        if (this.SecondaryImageHandler != null
            && this.SecondaryImageHandler.SupportedTypes.Contains(type)
            && this.SecondaryImageHandler.GetImageBounds(ms, type, out width, out height, out horizontalResolution, out verticalResolution))
        {
            return true;
        }

        if (this.TertiaryImageHandler != null
            && this.TertiaryImageHandler.SupportedTypes.Contains(type)
            && this.TertiaryImageHandler.GetImageBounds(ms, type, out width, out height, out horizontalResolution, out verticalResolution))
        {
            return true;
        }

        width = height = horizontalResolution = verticalResolution = 0;

        return false;
    }
}