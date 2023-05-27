/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// Handels themes in a package
/// </summary>
public class ExcelThemeManager
{
    ExcelWorkbook _wb;
    internal static string _defaultTheme="";
    internal ExcelThemeManager(ExcelWorkbook wb)
    {
        this._wb = wb;
    }
    ExcelTheme _theme = null;
    /// <summary>
    /// The current theme. Null if not theme exists.
    /// <seealso cref="CreateDefaultTheme"/>
    /// <seealso cref="Load(FileInfo)"/>
    /// <seealso cref="Load(Stream)"/>
    /// <seealso cref="Load(XmlDocument)"/>
    /// </summary>
    public ExcelTheme CurrentTheme
    {
        get
        {
            if(this._theme==null)
            {
                ZipPackageRelationshipCollection? rels = this._wb.Part.GetRelationshipsByType(ExcelPackage.schemaThemeRelationships);
                if (rels.Count>0)
                {
                    this._theme = new ExcelTheme(this._wb, rels.First());
                }
            }
            return this._theme;
        }
    }
    /// <summary>
    /// Create the default theme.
    /// </summary>
    public void CreateDefaultTheme()
    {
        if (this.CurrentTheme != null)
        {
            throw new InvalidOperationException("Can't create theme. Theme already exists");
        }

        if(string.IsNullOrEmpty(_defaultTheme))
        {
            _defaultTheme = StyleResourceManager.GetItem("DefaultTheme.Xml");
        }
        XmlDocument? themeXml = new XmlDocument();   
        themeXml.LoadXml(_defaultTheme);
        this.Load(themeXml);
    }
    internal ExcelTheme GetOrCreateTheme()
    {
        if(this.CurrentTheme==null)
        {
            this.CreateDefaultTheme();
        }
        return this._theme;
    }
    /// <summary>
    /// Delete the current theme
    /// </summary>
    public void DeleteCurrentTheme()
    {
        if(this.CurrentTheme==null)
        {
            return;
        }

        this._wb._package.ZipPackage.DeleteRelationship(this._theme.RelationshipId);
        this._wb._package.ZipPackage.DeletePart(this._theme.ThemeUri);
        this._theme = null;
    }
    /// <summary>
    /// Loads a .thmx file, exported from a Spread Sheet Application like Excel
    /// </summary>
    /// <param name="thmxFile">The path to the thmx file</param>
    public void Load(FileInfo thmxFile)
    {
        if(!thmxFile.Exists)
        {
            throw new FileNotFoundException($"{thmxFile.FullName} does not exist");
        }

        using MemoryStream? ms = RecyclableMemory.GetStream(File.ReadAllBytes(thmxFile.FullName));
        this.Load(ms);
    }
    /// <summary>
    /// Loads a theme XmlDocument. 
    /// Overwrites any previously set theme settings.
    /// </summary>
    /// <param name="themeXml">The theme xml</param>
    public void Load(XmlDocument themeXml)
    {
        this.DeleteCurrentTheme();
        if (this.CurrentTheme == null)
        {
            Uri? uri = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
            ZipPackagePart? part = this._wb._package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeTheme);
            themeXml.Save(part.GetStream());
            ZipPackageRelationship? rel = this._wb.Part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaThemeRelationships);
            this._theme = new ExcelTheme(this._wb, rel);
        }
    }
    /// <summary>
    /// Loads a .thmx file as a stream. Thmx files are exported from a Spread Sheet Application like Excel
    /// </summary>
    /// <param name="thmxStream">The thmx file as a stream</param>
    public void Load(Stream thmxStream)
    {
            
        ZipPackage p = new ZipPackage(thmxStream);
            
        ZipPackageRelationship? themeManagerRel = p.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").FirstOrDefault();
        if (themeManagerRel != null)
        {
            ZipPackagePart? themeManager = p.GetPart(themeManagerRel.TargetUri);
            ZipPackageRelationship? themeRel = themeManager.GetRelationshipsByType(ExcelPackage.schemaThemeRelationships).FirstOrDefault();
            if (themeRel != null)
            {
                ZipPackagePart? themePart = p.GetPart(UriHelper.ResolvePartUri(themeRel.SourceUri, themeRel.TargetUri));
                XmlDocument? themeXml = new XmlDocument();
                XmlHelper.LoadXmlSafe(themeXml, themePart.GetStream());
                this.Load(themeXml);
                foreach (ZipPackageRelationship? rel in themePart.GetRelationships())
                {   
                    ZipPackagePart? partToCopy = p.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                    Uri? uri = UriHelper.ResolvePartUri(this._theme.ThemeUri, rel.TargetUri);
                    ZipPackagePart? part = this._wb._package.ZipPackage.CreatePart(uri, partToCopy.ContentType);
                    Stream? stream = part.GetStream();
                    byte[]? b = ((MemoryStream)partToCopy.GetStream()).ToArray();
                    stream.Write(b, 0, b.Length);
                    stream.Flush();
                    this._theme.Part.CreateRelationship(uri, TargetMode.Internal, rel.RelationshipType);
                }

                this.SetNormalStyle();
            }
            else
            {
                throw new InvalidDataException("Thmx file is corrupt. Can't find theme part");
            }
        }
        else
        {
            throw new InvalidDataException("Thmx file is corrupt.");
        }
    }

    private void SetNormalStyle()
    {
        if (this._wb.Styles.NamedStyles.Count == 0)
        {
            return;
        }

        ExcelNamedStyleXml? style = this.GetNormalStyle();
        foreach(ExcelXfs? xfs in this._wb.Styles.CellXfs)
        {
            if (xfs.XfId == style.StyleXfId)
            {
                ExcelFontXml? font = this._wb.Styles.Fonts[xfs.FontId];
                font.Name = this.CurrentTheme.FontScheme.MinorFont[0].Typeface;
                font.Family = 2;
                font.Color.Theme = eThemeSchemeColor.Text1;
                font.Scheme = "minor";
            }
        }
    }

    private ExcelNamedStyleXml GetNormalStyle()
    {
        return this._wb.Styles.GetNormalStyle();
    }
    internal void Save()
    {
        if (this.CurrentTheme != null)
        {
            this._wb._package.SavePart(this.CurrentTheme.ThemeUri, this.CurrentTheme.ThemeXml);
        }
    }
}