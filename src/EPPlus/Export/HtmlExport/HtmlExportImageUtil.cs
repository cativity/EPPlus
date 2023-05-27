using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport;

internal static class HtmlExportImageUtil
{
    private static string GetClassName(string className, string optionalName)
    {
        if (string.IsNullOrEmpty(optionalName))
        {
            return optionalName;
        }

        className = className.Trim().Replace(" ", "-");
        string? newClassName = "";

        for (int i = 0; i < className.Length; i++)
        {
            char c = className[i];

            if (i == 0)
            {
                if (c == '-' || (c >= '0' && c <= '9'))
                {
                    newClassName = "_";

                    continue;
                }
            }

            if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c >= 0x00A0)
            {
                newClassName += c;
            }
        }

        return string.IsNullOrEmpty(newClassName) ? optionalName : newClassName;
    }

    internal static string GetPictureName(HtmlImage p)
    {
        string? hash = ((IPictureContainer)p.Picture).ImageHash;
        FileInfo? fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
        string? name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

        return GetClassName(name, hash);
    }

    internal static void AddImage(EpplusHtmlWriter writer, HtmlExportSettings settings, HtmlImage image, object value)
    {
        if (image != null)
        {
            string? name = GetPictureName(image);
            string imageName = GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
            writer.AddAttribute("alt", image.Picture.Name);

            if (settings.Pictures.AddNameAsId)
            {
                writer.AddAttribute("id", imageName);
            }

            writer.AddAttribute("class", $"{settings.StyleClassPrefix}image-{name} {settings.StyleClassPrefix}image-prop-{imageName}");
            writer.RenderBeginTag(HtmlElements.Img, true);
        }
    }
}