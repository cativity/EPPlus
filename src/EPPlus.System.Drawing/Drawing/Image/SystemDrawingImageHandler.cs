using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Image;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;

namespace OfficeOpenXml.SystemDrawing.Image;

public class SystemDrawingImageHandler : IImageHandler
{
    public SystemDrawingImageHandler()
    {
        if(IsWindows())
        {
            this.SupportedTypes= new HashSet<ePictureType>() { ePictureType.Bmp, ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Tif, ePictureType.Emf, ePictureType.Wmf };
        }
        else
        {
            this.SupportedTypes = new HashSet<ePictureType>() { ePictureType.Bmp, ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Tif };
        }
    }

    private static bool IsWindows()
    {
        if(Environment.OSVersion.Platform == PlatformID.Unix ||
#if(NET5_0_OR_GREATER)
           Environment.OSVersion.Platform == PlatformID.Other ||
#endif
           Environment.OSVersion.Platform == PlatformID.MacOSX)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    public HashSet<ePictureType> SupportedTypes
    {
        get;
    } 

    public Exception LastException { get; private set; }

    public bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution)
    {
        try
        {
            Bitmap? bmp = new Bitmap(image);
            width = bmp.Width;
            height = bmp.Height;
            horizontalResolution = bmp.HorizontalResolution;
            verticalResolution = bmp.VerticalResolution;
            return true;
        }
        catch(Exception ex)
        {
            width = 0;
            height = 0;
            horizontalResolution = 0;
            verticalResolution = 0;
            this.LastException = ex;
            return false;
        }
    }
    bool? _validForEnvironment = null;
    public bool ValidForEnvironment()
    {
        if (this._validForEnvironment.HasValue == false)
        {
            try
            {
                Graphics? g = Graphics.FromHwnd(IntPtr.Zero); //Fails if no gdi.
                this._validForEnvironment = true;
            }
            catch
            {
                this._validForEnvironment = false;
            }
        }
        return this._validForEnvironment.Value;
    }
}