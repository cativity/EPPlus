using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal static class ZipHelper
    {
        internal static string UncompressEntry(ZipInputStream zipStream, ZipEntry entry)
        {
            byte[]? content = new byte[entry.UncompressedSize];
            int size = zipStream.Read(content, 0, (int)entry.UncompressedSize);
            return Encoding.UTF8.GetString(content);
        }

        internal static ZipInputStream OpenZipResource()
        {
            Assembly? assembly = Assembly.GetExecutingAssembly();
            Stream? templateStream = assembly.GetManifestResourceStream("OfficeOpenXml.resources.DefaultTableStyles.cst");
            ZipInputStream? zipStream = new ZipInputStream(templateStream);
            return zipStream;
        }
    }
}
