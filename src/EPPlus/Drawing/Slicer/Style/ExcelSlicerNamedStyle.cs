/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer.Style
{
    /// <summary>
    /// A named table style that applies to tables only
    /// </summary>
    public class ExcelSlicerNamedStyle : XmlHelper
    {
        ExcelStyles _styles;
        internal Dictionary<eSlicerStyleElement, ExcelSlicerStyleElement> _dicSlicer = new Dictionary<eSlicerStyleElement, ExcelSlicerStyleElement>();
        internal Dictionary<eTableStyleElement, ExcelSlicerTableStyleElement> _dicTable = new Dictionary<eTableStyleElement, ExcelSlicerTableStyleElement>();
        XmlNode _tableStyleNode;
        internal ExcelSlicerNamedStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, XmlNode tableStyleNode, ExcelStyles styles) : base(nameSpaceManager, topNode)
        {
            this._styles = styles;
            if (tableStyleNode == null)
            {
                //TODO: Create table styles node with 
            }
            else
            {
                this._tableStyleNode = tableStyleNode;
                foreach (XmlNode node in tableStyleNode.ChildNodes)
                {
                    if (node is XmlElement e)
                    {
                        eTableStyleElement type = e.GetAttribute("type").ToEnum(eTableStyleElement.WholeTable);
                        this._dicTable.Add(type, new ExcelSlicerTableStyleElement(nameSpaceManager, node, styles, type));
                    }
                }
            }
            if (topNode.HasChildNodes)
            {
                foreach (XmlNode node in topNode?.FirstChild?.ChildNodes)
                {
                    if (node is XmlElement e)
                    {
                        eSlicerStyleElement type = e.GetAttribute("type").ToEnum(eSlicerStyleElement.SelectedItemWithData);
                        this._dicSlicer.Add(type, new ExcelSlicerStyleElement(nameSpaceManager, node, styles, type));
                    }
                }
            }
        }
        private ExcelSlicerTableStyleElement GetTableStyleElement(eTableStyleElement element)
        {
            if (this._dicTable.ContainsKey(element))
            {
                return this._dicTable[element];
            }
            ExcelSlicerTableStyleElement item;
            item = new ExcelSlicerTableStyleElement(this.NameSpaceManager, this._tableStyleNode, this._styles, element);
            this._dicTable.Add(element, item);
            return item;
        }
        private ExcelSlicerStyleElement GetSlicerStyleElement(eSlicerStyleElement element)
        {
            if (this._dicSlicer.ContainsKey(element))
            {
                return this._dicSlicer[element];
            }
            ExcelSlicerStyleElement item;
            item = new ExcelSlicerStyleElement(this.NameSpaceManager, this.TopNode, this._styles, element);
            this._dicSlicer.Add(element, item);
            return item;
        }

        /// <summary>
        /// The name of the table named style
        /// </summary>
        public string Name
        {
            get
            {
                return this.GetXmlNodeString("@name");
            }
            set
            {
                if (this._styles.SlicerStyles.ExistsKey(value) || this._styles.TableStyles.ExistsKey(value))
                {
                    throw new InvalidOperationException("Name already exists in the collection");
                }

                this.SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Applies to the entire content of a table or pivot table
        /// </summary>
        public ExcelSlicerTableStyleElement WholeTable
        {
            get
            {
                return this.GetTableStyleElement(eTableStyleElement.WholeTable);
            }
        }
        /// <summary>
        /// Applies to the header row of a table or pivot table
        /// </summary>
        public ExcelSlicerTableStyleElement HeaderRow
        {
            get
            {
                return this.GetTableStyleElement(eTableStyleElement.HeaderRow);
            }
        }
        /// <summary>
        /// Applies to slicer item that is selected
        /// </summary>
        public ExcelSlicerStyleElement SelectedItemWithData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.SelectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a select slicer item with no data.
        /// </summary>
        public ExcelSlicerStyleElement SelectedItemWithNoData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.SelectedItemWithNoData);
            }
        }

        /// <summary>
        /// Applies to a slicer item with data that is not selected
        /// </summary>
        public ExcelSlicerStyleElement UnselectedItemWithData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.UnselectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a slicer item with no data that is not selected
        /// </summary>
        public ExcelSlicerStyleElement UnselectedItemWithNoData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.UnselectedItemWithNoData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredSelectedItemWithData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.HoveredSelectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with no data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredSelectedItemWithNoData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.HoveredSelectedItemWithNoData);
            }
        }

        /// <summary>
        /// Applies to a slicer item with data that is not selected and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredUnselectedItemWithData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.HoveredUnselectedItemWithData);
            }
        }
        /// <summary>
        /// Applies to a selected slicer item with no data and over which the mouse is paused on
        /// </summary>
        public ExcelSlicerStyleElement HoveredUnselectedItemWithNoData
        {
            get
            {
                return this.GetSlicerStyleElement(eSlicerStyleElement.HoveredUnselectedItemWithNoData);
            }
        }

        internal void SetFromTemplate(ExcelSlicerNamedStyle templateStyle)
        {
            foreach (ExcelSlicerTableStyleElement? s in templateStyle._dicTable.Values)
            {
                ExcelSlicerTableStyleElement? element = this.GetTableStyleElement(s.Type);
                element.Style = (ExcelDxfSlicerStyle)s.Style.Clone();
            }
            foreach (ExcelSlicerStyleElement? s in templateStyle._dicSlicer.Values)
            {
                ExcelSlicerStyleElement? element = this.GetSlicerStyleElement(s.Type);
                element.Style = (ExcelDxfSlicerStyle)s.Style.Clone();
            }
        }
        internal void SetFromTemplate(eSlicerStyle templateStyle)
        {
            this.LoadTableTemplate("SlicerStyles", templateStyle.ToString());
        }
        private void LoadTableTemplate(string folder, string styleName)
        {
            ZipInputStream? zipStream = ZipHelper.OpenZipResource();
            ZipEntry entry;
            while ((entry = zipStream.GetNextEntry()) != null)
            {
                if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0 || !entry.FileName.StartsWith(folder))
                {
                    continue;
                }

                string? name = new FileInfo(entry.FileName).Name;
                name = name.Substring(0, name.Length - 4);
                if (name.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                {
                    string? xmlContent = ZipHelper.UncompressEntry(zipStream, entry);
                    XmlDocument? xml = new XmlDocument();
                    xml.LoadXml(xmlContent);

                    foreach (XmlElement elem in xml.DocumentElement.ChildNodes)
                    {
                        string? dxfXml = elem.InnerXml;
                        eTableStyleElement? tblType = elem.GetAttribute("name").ToEnum<eTableStyleElement>();
                        if(tblType==null)
                        {
                            eSlicerStyleElement? slicerType= elem.GetAttribute("name").ToEnum<eSlicerStyleElement>();
                            if(slicerType.HasValue)
                            {
                                ExcelSlicerStyleElement? se = this.GetSlicerStyleElement(slicerType.Value);
                                ExcelDxfSlicerStyle? dxf = new ExcelDxfSlicerStyle(this.NameSpaceManager, elem.FirstChild, this._styles, null);
                                se.Style = dxf;
                            }
                        }
                        else
                        {
                            ExcelSlicerTableStyleElement? te = this.GetTableStyleElement(tblType.Value);
                            ExcelDxfSlicerStyle? dxf = new ExcelDxfSlicerStyle(this.NameSpaceManager, elem.FirstChild, this._styles, null);
                            te.Style = dxf;
                        }
                    }
                }
            }
        }
    }
}
