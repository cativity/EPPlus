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
using System.Xml;
using System.IO;
using System.Globalization;
using OfficeOpenXml.Utils;
using System.Linq;
using OfficeOpenXml.Utils.Extensions;
using System.Collections.Generic;

namespace OfficeOpenXml;

/// <summary>
/// Provides access to the properties bag of the package
/// </summary>
public sealed class OfficeProperties : XmlHelper
{
    #region Private Properties

    private XmlDocument _xmlPropertiesCore;
    private XmlDocument _xmlPropertiesExtended;
    private XmlDocument _xmlPropertiesCustom;

    private Uri _uriPropertiesCore = new Uri("/docProps/core.xml", UriKind.Relative);
    private Uri _uriPropertiesExtended = new Uri("/docProps/app.xml", UriKind.Relative);
    private Uri _uriPropertiesCustom = new Uri("/docProps/custom.xml", UriKind.Relative);

    XmlHelper _coreHelper;
    XmlHelper _extendedHelper;
    XmlHelper _customHelper;
    private readonly Dictionary<string, XmlElement> _customProperties;
    private ExcelPackage _package;

    int _maxPid = 1;

    #endregion

    #region ExcelProperties Constructor

    /// <summary>
    /// Provides access to all the office document properties.
    /// </summary>
    /// <param name="package"></param>
    /// <param name="ns"></param>
    internal OfficeProperties(ExcelPackage package, XmlNamespaceManager ns)
        : base(ns)
    {
        this._package = package;

        this._coreHelper = XmlHelperFactory.Create(ns, this.CorePropertiesXml.SelectSingleNode("cp:coreProperties", this.NameSpaceManager));
        this._extendedHelper = XmlHelperFactory.Create(ns, this.ExtendedPropertiesXml);
        this._customHelper = XmlHelperFactory.Create(ns, this.CustomPropertiesXml);
        this._customProperties = new Dictionary<string, XmlElement>(StringComparer.CurrentCultureIgnoreCase);
        this.LoadCustomProperties();
    }

    private void LoadCustomProperties()
    {
        foreach (XmlElement node in this.CustomPropertiesXml.SelectNodes("ctp:Properties/ctp:property", this.NameSpaceManager))
        {
            this._customProperties.Add(node.GetAttribute("name"), node);

            if (ConvertUtil.TryParseIntString(node.GetAttribute("pid"), out int pid))
            {
                if (pid > this._maxPid)
                {
                    this._maxPid = pid;
                }
            }
        }
    }

    #endregion

    #region CorePropertiesXml

    /// <summary>
    /// Provides access to the XML document that holds all the code 
    /// document properties.
    /// </summary>
    public XmlDocument CorePropertiesXml
    {
        get
        {
            if (this._xmlPropertiesCore == null)
            {
                string xml =
                    string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>",
                                  ExcelPackage.schemaCore,
                                  ExcelPackage.schemaDc,
                                  ExcelPackage.schemaDcTerms,
                                  ExcelPackage.schemaDcmiType,
                                  ExcelPackage.schemaXsi);

                this._xmlPropertiesCore = this.GetXmlDocument(xml,
                                                              this._uriPropertiesCore,
                                                              @"application/vnd.openxmlformats-package.core-properties+xml",
                                                              @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
            }

            return this._xmlPropertiesCore;
        }
    }

    private XmlDocument GetXmlDocument(string startXml, Uri uri, string contentType, string relationship)
    {
        XmlDocument xmlDoc;

        if (this._package.ZipPackage.PartExists(uri))
        {
            xmlDoc = this._package.GetXmlFromUri(uri);
        }
        else
        {
            xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(startXml);

            // Create a the part and add to the package
            Packaging.ZipPackagePart part = this._package.ZipPackage.CreatePart(uri, contentType);

            // Save it to the package
            StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            xmlDoc.Save(stream);

            //stream.Close();
            Packaging.ZipPackage.Flush();

            // create the relationship between the workbook and the new shared strings part
            _ = this._package.ZipPackage.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), uri),
                                                        Packaging.TargetMode.Internal,
                                                        relationship);

            Packaging.ZipPackage.Flush();
        }

        return xmlDoc;
    }

    #endregion

    #region Core Properties

    const string TitlePath = "dc:title";

    /// <summary>
    /// Gets/sets the title property of the document (core property)
    /// </summary>
    public string Title
    {
        get { return this._coreHelper.GetXmlNodeString(TitlePath); }
        set { this._coreHelper.SetXmlNodeString(TitlePath, value); }
    }

    const string SubjectPath = "dc:subject";

    /// <summary>
    /// Gets/sets the subject property of the document (core property)
    /// </summary>
    public string Subject
    {
        get { return this._coreHelper.GetXmlNodeString(SubjectPath); }
        set { this._coreHelper.SetXmlNodeString(SubjectPath, value); }
    }

    const string AuthorPath = "dc:creator";

    /// <summary>
    /// Gets/sets the author property of the document (core property)
    /// </summary>
    public string Author
    {
        get { return this._coreHelper.GetXmlNodeString(AuthorPath); }
        set { this._coreHelper.SetXmlNodeString(AuthorPath, value); }
    }

    const string CommentsPath = "dc:description";

    /// <summary>
    /// Gets/sets the comments property of the document (core property)
    /// </summary>
    public string Comments
    {
        get { return this._coreHelper.GetXmlNodeString(CommentsPath); }
        set { this._coreHelper.SetXmlNodeString(CommentsPath, value); }
    }

    const string KeywordsPath = "cp:keywords";

    /// <summary>
    /// Gets/sets the keywords property of the document (core property)
    /// </summary>
    public string Keywords
    {
        get { return this._coreHelper.GetXmlNodeString(KeywordsPath); }
        set { this._coreHelper.SetXmlNodeString(KeywordsPath, value); }
    }

    const string LastModifiedByPath = "cp:lastModifiedBy";

    /// <summary>
    /// Gets/sets the lastModifiedBy property of the document (core property)
    /// </summary>
    public string LastModifiedBy
    {
        get { return this._coreHelper.GetXmlNodeString(LastModifiedByPath); }
        set { this._coreHelper.SetXmlNodeString(LastModifiedByPath, value); }
    }

    const string LastPrintedPath = "cp:lastPrinted";

    /// <summary>
    /// Gets/sets the lastPrinted property of the document (core property)
    /// </summary>
    public string LastPrinted
    {
        get { return this._coreHelper.GetXmlNodeString(LastPrintedPath); }
        set { this._coreHelper.SetXmlNodeString(LastPrintedPath, value); }
    }

    const string CreatedPath = "dcterms:created";

    /// <summary>
    /// Gets/sets the created property of the document (core property)
    /// </summary>
    public DateTime Created
    {
        get { return DateTime.TryParse(this._coreHelper.GetXmlNodeString(CreatedPath), out DateTime date) ? date : DateTime.MinValue; }
        set
        {
            string? dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
            this._coreHelper.SetXmlNodeString(CreatedPath, dateString);
            this._coreHelper.SetXmlNodeString(CreatedPath + "/@xsi:type", "dcterms:W3CDTF");
        }
    }

    const string CategoryPath = "cp:category";

    /// <summary>
    /// Gets/sets the category property of the document (core property)
    /// </summary>
    public string Category
    {
        get { return this._coreHelper.GetXmlNodeString(CategoryPath); }
        set { this._coreHelper.SetXmlNodeString(CategoryPath, value); }
    }

    const string ContentStatusPath = "cp:contentStatus";

    /// <summary>
    /// Gets/sets the status property of the document (core property)
    /// </summary>
    public string Status
    {
        get { return this._coreHelper.GetXmlNodeString(ContentStatusPath); }
        set { this._coreHelper.SetXmlNodeString(ContentStatusPath, value); }
    }

    #endregion

    #region Extended Properties

    #region ExtendedPropertiesXml

    /// <summary>
    /// Provides access to the XML document that holds the extended properties of the document (app.xml)
    /// </summary>
    public XmlDocument ExtendedPropertiesXml
    {
        get
        {
            return this._xmlPropertiesExtended ??=
                       this.GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                                                         ExcelPackage.schemaVt,
                                                         ExcelPackage.schemaExtended),
                                           this._uriPropertiesExtended,
                                           @"application/vnd.openxmlformats-officedocument.extended-properties+xml",
                                           @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        }
    }

    #endregion

    const string ApplicationPath = "xp:Properties/xp:Application";

    /// <summary>
    /// Gets/Set the Application property of the document (extended property)
    /// </summary>
    public string Application
    {
        get { return this._extendedHelper.GetXmlNodeString(ApplicationPath); }
        set { this._extendedHelper.SetXmlNodeString(ApplicationPath, value); }
    }

    const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";

    /// <summary>
    /// Gets/sets the HyperlinkBase property of the document (extended property)
    /// </summary>
    public Uri HyperlinkBase
    {
        get { return new Uri(this._extendedHelper.GetXmlNodeString(HyperlinkBasePath), UriKind.Absolute); }
        set { this._extendedHelper.SetXmlNodeString(HyperlinkBasePath, value.AbsoluteUri); }
    }

    const string AppVersionPath = "xp:Properties/xp:AppVersion";

    /// <summary>
    /// Gets/Set the AppVersion property of the document (extended property)
    /// </summary>
    public string AppVersion
    {
        get { return this._extendedHelper.GetXmlNodeString(AppVersionPath); }
        set
        {
            string[]? versions = value.Split('.');

            if (versions.Length != 2 || versions.Any(x => !x.IsInt()))
            {
                throw new ArgumentException("AppVersion should be in the format XX.YYYY. X and Y are numeric values");
            }

            this._extendedHelper.SetXmlNodeString(AppVersionPath, value);
        }
    }

    const string CompanyPath = "xp:Properties/xp:Company";

    /// <summary>
    /// Gets/sets the Company property of the document (extended property)
    /// </summary>
    public string Company
    {
        get { return this._extendedHelper.GetXmlNodeString(CompanyPath); }
        set { this._extendedHelper.SetXmlNodeString(CompanyPath, value); }
    }

    const string ManagerPath = "xp:Properties/xp:Manager";

    /// <summary>
    /// Gets/sets the Manager property of the document (extended property)
    /// </summary>
    public string Manager
    {
        get { return this._extendedHelper.GetXmlNodeString(ManagerPath); }
        set { this._extendedHelper.SetXmlNodeString(ManagerPath, value); }
    }

    const string ModifiedPath = "dcterms:modified";

    /// <summary>
    /// Gets/sets the modified property of the document (core property)
    /// </summary>
    public DateTime Modified
    {
        get { return DateTime.TryParse(this._coreHelper.GetXmlNodeString(ModifiedPath), out DateTime date) ? date : DateTime.MinValue; }
        set
        {
            string? dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
            this._coreHelper.SetXmlNodeString(ModifiedPath, dateString);
            this._coreHelper.SetXmlNodeString(ModifiedPath + "/@xsi:type", "dcterms:W3CDTF");
        }
    }

    const string LinksUpToDatePath = "xp:Properties/xp:LinksUpToDate";

    /// <summary>
    /// Indicates whether hyperlinks in a document are up-to-date
    /// </summary>
    public bool LinksUpToDate
    {
        get { return this._extendedHelper.GetXmlNodeBool(LinksUpToDatePath); }
        set { this._extendedHelper.SetXmlNodeBool(LinksUpToDatePath, value); }
    }

    const string HyperlinksChangedPath = "xp:Properties/xp:HyperlinksChanged";

    /// <summary>
    /// Hyperlinks need update
    /// </summary>
    public bool HyperlinksChanged
    {
        get { return this._extendedHelper.GetXmlNodeBool(HyperlinksChangedPath); }
        set { this._extendedHelper.SetXmlNodeBool(HyperlinksChangedPath, value); }
    }

    const string ScaleCropPath = "xp:Properties/xp:ScaleCrop";

    /// <summary>
    /// Display mode of the document thumbnail. True to enable scaling. False to enable cropping.
    /// </summary>
    public bool ScaleCrop
    {
        get { return this._extendedHelper.GetXmlNodeBool(ScaleCropPath); }
        set { this._extendedHelper.SetXmlNodeBool(ScaleCropPath, value); }
    }

    const string SharedDocPath = "xp:Properties/xp:SharedDoc";

    /// <summary>
    /// If true, document is shared between multiple producers.
    /// </summary>
    public bool SharedDoc
    {
        get { return this._extendedHelper.GetXmlNodeBool(SharedDocPath); }
        set { this._extendedHelper.SetXmlNodeBool(SharedDocPath, value); }
    }

    #region Get and Set Extended Properties

    /// <summary>
    /// Get the value of an extended property 
    /// </summary>
    /// <param name="propertyName">The name of the property</param>
    /// <returns>The value</returns>
    public string GetExtendedPropertyValue(string propertyName)
    {
        string retValue = null;
        string searchString = string.Format("xp:Properties/xp:{0}", propertyName);
        XmlNode node = this.ExtendedPropertiesXml.SelectSingleNode(searchString, this.NameSpaceManager);

        if (node != null)
        {
            retValue = node.InnerText;
        }

        return retValue;
    }

    /// <summary>
    /// Set the value for an extended property
    /// </summary>
    /// <param name="propertyName">The name of the property</param>
    /// <param name="value">The value</param>
    public void SetExtendedPropertyValue(string propertyName, string value)
    {
        string propertyPath = string.Format("xp:Properties/xp:{0}", propertyName);
        this._extendedHelper.SetXmlNodeString(propertyPath, value);
    }

    #endregion

    #endregion

    #region Custom Properties

    #region CustomPropertiesXml

    /// <summary>
    /// Provides access to the XML document which holds the document's custom properties
    /// </summary>
    public XmlDocument CustomPropertiesXml
    {
        get
        {
            return this._xmlPropertiesCustom ??=
                       this.GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                                                         ExcelPackage.schemaVt,
                                                         ExcelPackage.schemaCustom),
                                           this._uriPropertiesCustom,
                                           @"application/vnd.openxmlformats-officedocument.custom-properties+xml",
                                           @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
        }
    }

    #endregion

    #region Get and Set Custom Properties

    /// <summary>
    /// Gets the value of a custom property
    /// </summary>
    /// <param name="propertyName">The name of the property</param>
    /// <returns>The current value of the property</returns>
    public object GetCustomPropertyValue(string propertyName)
    {
        if (this._customProperties.ContainsKey(propertyName))
        {
            XmlElement? node = this._customProperties[propertyName];
            string value = node.LastChild.InnerText;

            switch (node.LastChild.LocalName)
            {
                case "filetime":
                    if (DateTime.TryParse(value, out DateTime dt))
                    {
                        return dt;
                    }
                    else
                    {
                        return null;
                    }

                case "i4":
                    if (int.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out int i))
                    {
                        return i;
                    }
                    else
                    {
                        return null;
                    }

                case "r8":
                    if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double d))
                    {
                        return d;
                    }
                    else
                    {
                        return null;
                    }

                case "bool":
                    if (value == "true")
                    {
                        return true;
                    }
                    else if (value == "false")
                    {
                        return false;
                    }
                    else
                    {
                        return null;
                    }

                default:
                    return value;
            }
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Allows you to set the value of a current custom property or create your own custom property.  
    /// </summary>
    /// <param name="propertyName">The name of the property</param>
    /// <param name="value">The value of the property</param>
    public void SetCustomPropertyValue(string propertyName, object value)
    {
        XmlNode allProps = this.CustomPropertiesXml.SelectSingleNode(@"ctp:Properties", this.NameSpaceManager);
        XmlElement node;

        if (this._customProperties.ContainsKey(propertyName))
        {
            node = this._customProperties[propertyName];
            node.IsEmpty = true;
        }
        else
        {
            node = this.CustomPropertiesXml.CreateElement("property", ExcelPackage.schemaCustom);
            node.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
            this._maxPid++;
            node.SetAttribute("pid", this._maxPid.ToString()); // custom property pid
            node.SetAttribute("name", propertyName);

            this._customProperties.Add(propertyName, node);
            _ = allProps.AppendChild(node);
        }

        XmlElement valueElem;

        if (value is bool)
        {
            valueElem = this.CustomPropertiesXml.CreateElement("vt", "bool", ExcelPackage.schemaVt);
            valueElem.InnerText = value.ToString().ToLower(CultureInfo.InvariantCulture);
        }
        else if (value is DateTime)
        {
            valueElem = this.CustomPropertiesXml.CreateElement("vt", "filetime", ExcelPackage.schemaVt);
            valueElem.InnerText = ((DateTime)value).AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
        }
        else if (value is short || value is int)
        {
            valueElem = this.CustomPropertiesXml.CreateElement("vt", "i4", ExcelPackage.schemaVt);
            valueElem.InnerText = value.ToString();
        }
        else if (value is double || value is decimal || value is float || value is long)
        {
            valueElem = this.CustomPropertiesXml.CreateElement("vt", "r8", ExcelPackage.schemaVt);

            if (value is double)
            {
                valueElem.InnerText = ((double)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is float)
            {
                valueElem.InnerText = ((float)value).ToString(CultureInfo.InvariantCulture);
            }
            else if (value is decimal)
            {
                valueElem.InnerText = ((decimal)value).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                valueElem.InnerText = value.ToString();
            }
        }
        else
        {
            valueElem = this.CustomPropertiesXml.CreateElement("vt", "lpwstr", ExcelPackage.schemaVt);
            valueElem.InnerText = value.ToString();
        }

        _ = node.AppendChild(valueElem);
    }

    #endregion

    #endregion

    #region Save

    /// <summary>
    /// Saves the document properties back to the package.
    /// </summary>
    internal void Save()
    {
        if (this._xmlPropertiesCore != null)
        {
            this._package.SavePart(this._uriPropertiesCore, this._xmlPropertiesCore);
        }

        if (this._xmlPropertiesExtended != null)
        {
            this._package.SavePart(this._uriPropertiesExtended, this._xmlPropertiesExtended);
        }

        if (this._xmlPropertiesCustom != null)
        {
            this._package.SavePart(this._uriPropertiesCustom, this._xmlPropertiesCustom);
        }
    }

    #endregion
}