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

using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Theme;

/// <summary>
/// The base class for a theme
/// </summary>
public class ExcelThemeBase : XmlHelper, IPictureRelationDocument
{
    readonly string _colorSchemePath = "{0}a:clrScheme";
    readonly string _fontSchemePath = "{0}a:fontScheme";
    readonly string _fmtSchemePath = "{0}a:fmtScheme";
    readonly ExcelPackage _pck;
    Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();

    internal ExcelThemeBase(ExcelPackage package, XmlNamespaceManager nsm, ZipPackageRelationship rel, string path)
        : base(nsm, null)
    {
        this.ThemeUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
        this.Part = package.ZipPackage.GetPart(this.ThemeUri);
        this.RelationshipId = rel.Id;
        this.ThemeXml = new XmlDocument();
        LoadXmlSafe(this.ThemeXml, this.Part.GetStream());
        this.TopNode = this.ThemeXml.DocumentElement;

        this._colorSchemePath = string.Format(this._colorSchemePath, path);
        this._fontSchemePath = string.Format(this._fontSchemePath, path);
        this._fmtSchemePath = string.Format(this._fmtSchemePath, path);
        this._pck = package;

        if (!this.NameSpaceManager.HasNamespace("a"))
        {
            this.NameSpaceManager.AddNamespace("a", ExcelPackage.schemaDrawings);
        }
    }

    internal Uri ThemeUri { get; set; }

    internal ZipPackagePart Part { get; set; }

    /// <summary>
    /// The Theme Xml
    /// </summary>
    public XmlDocument ThemeXml { get; internal set; }

    internal string RelationshipId { get; set; }

    internal ExcelColorScheme _colorScheme;

    /// <summary>
    /// Defines the color scheme
    /// </summary>
    public ExcelColorScheme ColorScheme
    {
        get
        {
            return this._colorScheme ??=
                           new ExcelColorScheme(this.NameSpaceManager, this.TopNode.SelectSingleNode(this._colorSchemePath, this.NameSpaceManager));
        }
    }

    internal ExcelFontScheme _fontScheme;

    /// <summary>
    /// Defines the font scheme
    /// </summary>
    public ExcelFontScheme FontScheme
    {
        get
        {
            return this._fontScheme ??= new ExcelFontScheme(this._pck,
                                                              this.NameSpaceManager,
                                                              this.TopNode.SelectSingleNode(this._fontSchemePath, this.NameSpaceManager));
        }
    }

    private ExcelFormatScheme _formatScheme;

    /// <summary>
    /// The background fill styles, effect styles, fill styles, and line styles which define the style matrix for a theme
    /// </summary>
    public ExcelFormatScheme FormatScheme
    {
        get
        {
            return this._formatScheme ??= new ExcelFormatScheme(this.NameSpaceManager,
                                                                  this.TopNode.SelectSingleNode(this._fmtSchemePath, this.NameSpaceManager),
                                                                  this);
        }
    }

    ExcelPackage IPictureRelationDocument.Package
    {
        get => this._pck;
    }

    Dictionary<string, HashInfo> IPictureRelationDocument.Hashes
    {
        get => this._hashes;
    }

    ZipPackagePart IPictureRelationDocument.RelatedPart
    {
        get => this.Part;
    }

    Uri IPictureRelationDocument.RelatedUri
    {
        get => this.ThemeUri;
    }
}