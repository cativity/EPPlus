/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// Base class for table and pivot table slicer caches
    /// </summary>
    public abstract class ExcelSlicerCache : XmlHelper
    {
        internal ExcelSlicerCache(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
        }
        internal void CreatePart(ExcelWorkbook wb)
        {
            ZipPackage? p = wb._package.ZipPackage;
            this.Uri = GetNewUri(p, "/xl/slicerCaches/slicerCache{0}.xml");
            this.Part = p.CreatePart(this.Uri, "application/vnd.ms-excel.slicerCache+xml");
            this.CacheRel = wb.Part.CreateRelationship(UriHelper.GetRelativeUri(wb.WorkbookUri, this.Uri), TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicerCache);
            this.SlicerCacheXml = new XmlDocument();
            this.SlicerCacheXml.LoadXml(GetStartXml());
        }
        internal ZipPackageRelationship CacheRel{ get; set; }
        internal ZipPackagePart Part { get; set; }
        internal Uri Uri { get; set; }
        /// <summary>
        /// The slicer cache xml document
        /// </summary>
        public XmlDocument SlicerCacheXml { get; protected internal set; }
        /// <summary>
        /// The name of the slicer cache
        /// </summary>
        public string Name
        {
            get
            {
                return this.GetXmlNodeString("@name");
            }
            internal protected set
            {
                this.SetXmlNodeString("@name",value);
            }
        }
        /// <summary>
        /// The name of the source field or column.
        /// </summary>
        public string SourceName
        {
            get
            {
                return this.GetXmlNodeString("@sourceName");
            }
            internal protected set
            {
                this.SetXmlNodeString("@sourceName", value);
            }
        }
        /// <summary>
        /// The source of the slicer.
        /// </summary>
        public abstract eSlicerSourceType SourceType
        {
            get;
        }

        internal abstract void Init(ExcelWorkbook wb);

        internal static string GetStartXml()
        {
            return $"<slicerCacheDefinition sourceName=\"\" xr10:uid=\"{{{Guid.NewGuid()}}}\" name=\"\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" mc:Ignorable=\"x xr10\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" />";
        }
        internal void CreateWorkbookReference(ExcelWorkbook wb, string uriGuid)
        {
            wb.Names.AddFormula(this.Name, "#N/A");
            if(!wb.SlicerCaches.ContainsKey(this.Name))
            {
                wb.SlicerCaches.Add(this.Name, this);
            }

            string prefix;
            if(this.GetType()==typeof(ExcelPivotTableSlicerCache))
            {
                prefix = "x14";
            }
            else
            {
                prefix = "x15";
            }
            XmlNode? extNode = wb.GetOrCreateExtLstSubNode(uriGuid, prefix, new string[] { ExtLstUris.WorkbookSlicerPivotTableUri, ExtLstUris.WorkbookSlicerTableUri });
            if (extNode.InnerXml=="")
            {
                extNode.InnerXml = $"<{prefix}:slicerCaches />";
            }
            XmlNode? slNode = extNode.FirstChild;
            XmlHelper? xh = XmlHelperFactory.Create(this.NameSpaceManager, slNode);
            XmlElement? element = (XmlElement)xh.CreateNode("x14:slicerCache", false, true);
            element.SetAttribute("id", ExcelPackage.schemaRelationships, this.CacheRel.Id);
        }
    }
}
