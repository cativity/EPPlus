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

namespace OfficeOpenXml
{
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
        internal OfficeProperties(ExcelPackage package, XmlNamespaceManager ns) :
            base(ns)
        {
            _package = package;

            _coreHelper = XmlHelperFactory.Create(ns, CorePropertiesXml.SelectSingleNode("cp:coreProperties", NameSpaceManager));
            _extendedHelper = XmlHelperFactory.Create(ns, ExtendedPropertiesXml);
            _customHelper = XmlHelperFactory.Create(ns, CustomPropertiesXml);
            _customProperties = new Dictionary<string, XmlElement>(StringComparer.CurrentCultureIgnoreCase);
            LoadCustomProperties();
        }

        private void LoadCustomProperties()
        {
            foreach (XmlElement node in CustomPropertiesXml.SelectNodes("ctp:Properties/ctp:property", NameSpaceManager))
            {
                _customProperties.Add(node.GetAttribute("name"), node);
                if (Utils.ConvertUtil.TryParseIntString(node.GetAttribute("pid"), out int pid))
                {
                    if (pid > _maxPid)
                    {
                        _maxPid = pid;
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
                if (_xmlPropertiesCore == null)
                {
                    string xml = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>",
                        ExcelPackage.schemaCore,
                        ExcelPackage.schemaDc,
                        ExcelPackage.schemaDcTerms,
                        ExcelPackage.schemaDcmiType,
                        ExcelPackage.schemaXsi);

                    _xmlPropertiesCore = GetXmlDocument(xml, _uriPropertiesCore, @"application/vnd.openxmlformats-package.core-properties+xml", @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                }
                return (_xmlPropertiesCore);
            }
        }

        private XmlDocument GetXmlDocument(string startXml, Uri uri, string contentType, string relationship)
        {
            XmlDocument xmlDoc;
            if (_package.ZipPackage.PartExists(uri))
            {
                xmlDoc = this._package.GetXmlFromUri(uri);
            }
            else
            {
                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(startXml);

                // Create a the part and add to the package
                Packaging.ZipPackagePart part = _package.ZipPackage.CreatePart(uri, contentType);

                // Save it to the package
                StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                xmlDoc.Save(stream);
                //stream.Close();
                _package.ZipPackage.Flush();

                // create the relationship between the workbook and the new shared strings part
                _package.ZipPackage.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), uri), Packaging.TargetMode.Internal, relationship);
                _package.ZipPackage.Flush();
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
            get { return _coreHelper.GetXmlNodeString(TitlePath); }
            set { _coreHelper.SetXmlNodeString(TitlePath, value); }
        }

        const string SubjectPath = "dc:subject";
        /// <summary>
        /// Gets/sets the subject property of the document (core property)
        /// </summary>
        public string Subject
        {
            get { return _coreHelper.GetXmlNodeString(SubjectPath); }
            set { _coreHelper.SetXmlNodeString(SubjectPath, value); }
        }

        const string AuthorPath = "dc:creator";
        /// <summary>
        /// Gets/sets the author property of the document (core property)
        /// </summary>
        public string Author
        {
            get { return _coreHelper.GetXmlNodeString(AuthorPath); }
            set { _coreHelper.SetXmlNodeString(AuthorPath, value); }
        }

        const string CommentsPath = "dc:description";
        /// <summary>
        /// Gets/sets the comments property of the document (core property)
        /// </summary>
        public string Comments
        {
            get { return _coreHelper.GetXmlNodeString(CommentsPath); }
            set { _coreHelper.SetXmlNodeString(CommentsPath, value); }
        }

        const string KeywordsPath = "cp:keywords";
        /// <summary>
        /// Gets/sets the keywords property of the document (core property)
        /// </summary>
        public string Keywords
        {
            get { return _coreHelper.GetXmlNodeString(KeywordsPath); }
            set { _coreHelper.SetXmlNodeString(KeywordsPath, value); }
        }

        const string LastModifiedByPath = "cp:lastModifiedBy";
        /// <summary>
        /// Gets/sets the lastModifiedBy property of the document (core property)
        /// </summary>
        public string LastModifiedBy
        {
            get { return _coreHelper.GetXmlNodeString(LastModifiedByPath); }
            set
            {
                _coreHelper.SetXmlNodeString(LastModifiedByPath, value);
            }
        }

        const string LastPrintedPath = "cp:lastPrinted";
        /// <summary>
        /// Gets/sets the lastPrinted property of the document (core property)
        /// </summary>
        public string LastPrinted
        {
            get { return _coreHelper.GetXmlNodeString(LastPrintedPath); }
            set { _coreHelper.SetXmlNodeString(LastPrintedPath, value); }
        }

        const string CreatedPath = "dcterms:created";

        /// <summary>
	    /// Gets/sets the created property of the document (core property)
	    /// </summary>
	    public DateTime Created
	    {
	        get
	        {
	            DateTime date;
	            return DateTime.TryParse(_coreHelper.GetXmlNodeString(CreatedPath), out date) ? date : DateTime.MinValue;
	        }
	        set
	        {
	            string? dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
	            _coreHelper.SetXmlNodeString(CreatedPath, dateString);
                _coreHelper.SetXmlNodeString(CreatedPath + "/@xsi:type", "dcterms:W3CDTF");
	        }
	    }

        const string CategoryPath = "cp:category";
        /// <summary>
        /// Gets/sets the category property of the document (core property)
        /// </summary>
        public string Category
        {
            get { return _coreHelper.GetXmlNodeString(CategoryPath); }
            set { _coreHelper.SetXmlNodeString(CategoryPath, value); }
        }

        const string ContentStatusPath = "cp:contentStatus";
        /// <summary>
        /// Gets/sets the status property of the document (core property)
        /// </summary>
        public string Status
        {
            get { return _coreHelper.GetXmlNodeString(ContentStatusPath); }
            set { _coreHelper.SetXmlNodeString(ContentStatusPath, value); }
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
                if (_xmlPropertiesExtended == null)
                {
                    _xmlPropertiesExtended = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaExtended),
                        _uriPropertiesExtended,
                        @"application/vnd.openxmlformats-officedocument.extended-properties+xml",
                        @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                }
                return (_xmlPropertiesExtended);
            }
        }
        #endregion

        const string ApplicationPath = "xp:Properties/xp:Application";
        /// <summary>
        /// Gets/Set the Application property of the document (extended property)
        /// </summary>
        public string Application
        {
            get { return _extendedHelper.GetXmlNodeString(ApplicationPath); }
            set { _extendedHelper.SetXmlNodeString(ApplicationPath, value); }
        }

        const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";
        /// <summary>
        /// Gets/sets the HyperlinkBase property of the document (extended property)
        /// </summary>
        public Uri HyperlinkBase
        {
            get { return new Uri(_extendedHelper.GetXmlNodeString(HyperlinkBasePath), UriKind.Absolute); }
            set { _extendedHelper.SetXmlNodeString(HyperlinkBasePath, value.AbsoluteUri); }
        }

        const string AppVersionPath = "xp:Properties/xp:AppVersion";
        /// <summary>
        /// Gets/Set the AppVersion property of the document (extended property)
        /// </summary>
        public string AppVersion
        {
            get { return _extendedHelper.GetXmlNodeString(AppVersionPath); }
            set 
            {
                string[]? versions = value.Split('.');
                if(versions.Length!=2 || versions.Any(x=>!x.IsInt()))
                {
                    throw (new ArgumentException("AppVersion should be in the format XX.YYYY. X and Y are numeric values"));
                }
                _extendedHelper.SetXmlNodeString(AppVersionPath, value); 
            }
        }
        const string CompanyPath = "xp:Properties/xp:Company";

        /// <summary>
        /// Gets/sets the Company property of the document (extended property)
        /// </summary>
        public string Company
        {
            get { return _extendedHelper.GetXmlNodeString(CompanyPath); }
            set { _extendedHelper.SetXmlNodeString(CompanyPath, value); }
        }

        const string ManagerPath = "xp:Properties/xp:Manager";
        /// <summary>
        /// Gets/sets the Manager property of the document (extended property)
        /// </summary>
        public string Manager
        {
            get { return _extendedHelper.GetXmlNodeString(ManagerPath); }
            set { _extendedHelper.SetXmlNodeString(ManagerPath, value); }
        }

        const string ModifiedPath = "dcterms:modified";
	    /// <summary>
	    /// Gets/sets the modified property of the document (core property)
	    /// </summary>
	    public DateTime Modified
	    {
	        get
	        {
	            DateTime date;
	            return DateTime.TryParse(_coreHelper.GetXmlNodeString(ModifiedPath), out date) ? date : DateTime.MinValue;
	        }
	        set
	        {
	            string? dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
	            _coreHelper.SetXmlNodeString(ModifiedPath, dateString);
                _coreHelper.SetXmlNodeString(ModifiedPath + "/@xsi:type", "dcterms:W3CDTF");
	        }
	    }
        const string LinksUpToDatePath = "xp:Properties/xp:LinksUpToDate";
        /// <summary>
        /// Indicates whether hyperlinks in a document are up-to-date
        /// </summary>
        public bool LinksUpToDate
        {
            get { return _extendedHelper.GetXmlNodeBool(LinksUpToDatePath); }
            set { _extendedHelper.SetXmlNodeBool(LinksUpToDatePath, value); }
        }
        const string HyperlinksChangedPath = "xp:Properties/xp:HyperlinksChanged";
        /// <summary>
        /// Hyperlinks need update
        /// </summary>
        public bool HyperlinksChanged
        {
            get { return _extendedHelper.GetXmlNodeBool(HyperlinksChangedPath); }
            set { _extendedHelper.SetXmlNodeBool(HyperlinksChangedPath, value); }
        }
        const string ScaleCropPath = "xp:Properties/xp:ScaleCrop";
        /// <summary>
        /// Display mode of the document thumbnail. True to enable scaling. False to enable cropping.
        /// </summary>
        public bool ScaleCrop
        {
            get { return _extendedHelper.GetXmlNodeBool(ScaleCropPath); }
            set { _extendedHelper.SetXmlNodeBool(ScaleCropPath, value); }
        }


        const string SharedDocPath = "xp:Properties/xp:SharedDoc";
        /// <summary>
        /// If true, document is shared between multiple producers.
        /// </summary>
        public bool SharedDoc
        {
            get { return _extendedHelper.GetXmlNodeBool(SharedDocPath); }
            set { _extendedHelper.SetXmlNodeBool(SharedDocPath, value); }
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
            XmlNode node = ExtendedPropertiesXml.SelectSingleNode(searchString, NameSpaceManager);
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
        public void SetExtendedPropertyValue(string propertyName, string value){
            string propertyPath = string.Format("xp:Properties/xp:{0}", propertyName);
            _extendedHelper.SetXmlNodeString(propertyPath, value);
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
                if (_xmlPropertiesCustom == null)
                {
                    _xmlPropertiesCustom = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaCustom),
                         _uriPropertiesCustom, 
                         @"application/vnd.openxmlformats-officedocument.custom-properties+xml",
                         @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                }
                return (_xmlPropertiesCustom);
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
            if (_customProperties.ContainsKey(propertyName))
            {
                XmlElement? node = _customProperties[propertyName];
                string value = node.LastChild.InnerText;
                switch (node.LastChild.LocalName)
                {
                    case "filetime":
                        DateTime dt;
                        if (DateTime.TryParse(value, out dt))
                        {
                            return dt;
                        }
                        else
                        {
                            return null;
                        }
                    case "i4":
                        int i;
                        if (int.TryParse(value, System.Globalization.NumberStyles.Number, CultureInfo.InvariantCulture, out i))
                        {
                            return i;
                        }
                        else
                        {
                            return null;
                        }
                    case "r8":
                        double d;
                        if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out d))
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
            XmlNode allProps = CustomPropertiesXml.SelectSingleNode(@"ctp:Properties", NameSpaceManager);
            XmlElement node;
            if (_customProperties.ContainsKey(propertyName))
            {
                node = _customProperties[propertyName];
                node.IsEmpty=true;
            }
            else
            {                
                node = CustomPropertiesXml.CreateElement("property", ExcelPackage.schemaCustom);
                node.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                _maxPid++;
                node.SetAttribute("pid", _maxPid.ToString());  // custom property pid
                node.SetAttribute("name", propertyName);

                _customProperties.Add(propertyName, node);
                allProps.AppendChild(node);
            }
            XmlElement valueElem;
            if (value is bool)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "bool", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString().ToLower(CultureInfo.InvariantCulture);
            }
            else if (value is DateTime)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "filetime", ExcelPackage.schemaVt);
                valueElem.InnerText = ((DateTime)value).AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
            }
            else if (value is short || value is int)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "i4", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString();
            }
            else if (value is double || value is decimal || value is float || value is long)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "r8", ExcelPackage.schemaVt);
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
                valueElem = CustomPropertiesXml.CreateElement("vt", "lpwstr", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString();
            }
            node.AppendChild(valueElem);
        }
        #endregion
        #endregion

        #region Save
        /// <summary>
        /// Saves the document properties back to the package.
        /// </summary>
        internal void Save()
        {
            if (_xmlPropertiesCore != null)
            {
                _package.SavePart(_uriPropertiesCore, _xmlPropertiesCore);
                }
            if (_xmlPropertiesExtended != null)
            {
                _package.SavePart(_uriPropertiesExtended, _xmlPropertiesExtended);
            }
            if (_xmlPropertiesCustom != null)
            {
                _package.SavePart(_uriPropertiesCustom, _xmlPropertiesCustom);
            }

        }
        #endregion

    }
}
