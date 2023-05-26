/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/28/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Base class for external references
    /// </summary>
    public abstract class ExcelExternalLink
    {
        internal ExcelWorkbook _wb;
        internal ExcelExternalLink(ExcelWorkbook wb)
        {
            this._wb = wb;
            this.As = new ExcelExternalLinkAsType(this);
            this.Part = null;
            this.WorkbookElement = null;
        }
        internal ExcelExternalLink(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement)
        {
            this._wb = wb;
            this.As = new ExcelExternalLinkAsType(this);
            this.Part = part;
            this.WorkbookElement = workbookElement;
        }
        /// <summary>
        /// The type of external link
        /// </summary>
        public abstract eExternalLinkType ExternalLinkType
        {
            get;
        }
        internal abstract void Save(StreamWriter sw);
        internal XmlElement WorkbookElement
        {
            get;
            set;
        }

        internal ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides an easy way to type cast the object to it's top level class
        /// </summary>
        public ExcelExternalLinkAsType As
        {
            get;
        }
        /// <summary>
        /// Returns the string representation of the object.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return this.ExternalLinkType.ToString();
        }
        /// <summary>
        /// The index of the external link. The index can be used in formulas between brackets to reference this link.
        /// </summary>
        /// <example>
        /// <code>worksheet.Cells["A1"].Formula="'[1]Sheet1'!A1"</code>
        /// </example>
        public int Index
        {
            get
            {
                return this._wb.ExternalLinks.GetIndex(this)+1;
            }
        }
        internal static bool HasWebProtocol(string uriPath)
        {
            return uriPath.StartsWith("http:") || uriPath.StartsWith("https:") || uriPath.StartsWith("ftp:") || uriPath.StartsWith("ftps:");
        }
        internal List<string> _errors = new List<string>();
        /// <summary>
        /// A list of errors that occured during load or update of the external workbook.
        /// </summary>
        public List<string> ErrorLog
        {
            get
            {
                return this._errors;
            }
        }

    }
}