﻿/*************************************************************************************************
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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public abstract class ExcelSlicerCache : XmlHelper
    {
        internal ExcelSlicerCache(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
        }
        internal ZipPackageRelationship CacheRel{ get; set; }
        internal ZipPackagePart Part { get; set; }
        public XmlDocument SlicerCacheXml { get; }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
        }
        public string SourceName
        {
            get
            {
                return GetXmlNodeString("@sourceName");
            }
        }
        public abstract eSlicerSourceType SourceType
        {
            get;
        }

        internal abstract void Init(ExcelWorkbook wb);
    }
}