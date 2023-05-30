﻿/*************************************************************************************************
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

using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

internal class ExcelSlicerXmlSources : XmlHelper
{
    const string _tableUId = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";
    const string _pivotTableUId = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";
    internal List<ExcelSlicerXmlSource> _list = new List<ExcelSlicerXmlSource>();
    internal ZipPackagePart _part;

    internal ExcelSlicerXmlSources(XmlNamespaceManager nsm, XmlNode topNode, ZipPackagePart part)
        : base(nsm, topNode)
    {
        this._part = part;

        foreach (XmlNode node in this.GetNodes("d:extLst/d:ext"))
        {
            switch (node.Attributes["uri"].Value)
            {
                case _tableUId: //Table slicer
                    foreach (XmlNode slicerNode in node.SelectNodes("x14:slicerList/x14:slicer", this.NameSpaceManager))
                    {
                        this._list.Add(new ExcelSlicerXmlSource(eSlicerSourceType.Table, part, slicerNode.Attributes["r:id"].Value));
                    }

                    break;

                case _pivotTableUId: //Pivot table slicer
                    foreach (XmlNode slicerNode in node.SelectNodes("x14:slicerList/x14:slicer", this.NameSpaceManager))
                    {
                        this._list.Add(new ExcelSlicerXmlSource(eSlicerSourceType.PivotTable, part, slicerNode.Attributes["r:id"].Value));
                    }

                    break;

                default:
                    break;
            }
        }
    }

    internal ExcelSlicerXmlSource GetOrCreateSource(eSlicerSourceType sourceType)
    {
        ExcelSlicerXmlSource? src = this.GetSources(sourceType).FirstOrDefault();

        if (src == null)
        {
            switch (sourceType)
            {
                case eSlicerSourceType.Table:
                    src = new ExcelSlicerXmlSource(eSlicerSourceType.Table, this._part, null);
                    this._list.Add(src);

                    break;

                case eSlicerSourceType.PivotTable:
                    src = new ExcelSlicerXmlSource(eSlicerSourceType.PivotTable, this._part, null);
                    this._list.Add(src);

                    break;
            }
        }

        return src;
    }

    internal XmlNode GetSource(string name, eSlicerSourceType sourceType, out ExcelSlicerXmlSource source)
    {
        foreach (ExcelSlicerXmlSource? s in this.GetSources(sourceType))
        {
            XmlNode? n = s.XmlDocument.DocumentElement.SelectSingleNode($"x14:slicer[@name=\"{name}\"]", this.NameSpaceManager);

            if (n != null)
            {
                source = s;

                return n;
            }
        }

        source = null;

        return null;
    }

    private IEnumerable<ExcelSlicerXmlSource> GetSources(eSlicerSourceType sourceType) => this._list.Where(x => x.Type == sourceType);

    internal void Save()
    {
        foreach (ExcelSlicerXmlSource? xs in this._list)
        {
            StreamWriter? stream = new StreamWriter(xs.Part.GetStream(FileMode.Create, FileAccess.Write));
            xs.XmlDocument.Save(stream);
        }
    }

    internal void Remove(ExcelSlicerXmlSource source)
    {
        _ = this._list.Remove(source);
        this._part.Package.DeletePart(source.Uri);
    }
}

internal class ExcelSlicerXmlSource : ExcelXmlSource
{
    internal ExcelSlicerXmlSource(eSlicerSourceType type, ZipPackagePart relPart, string relId)
        : base(relPart, relId) =>
        this.Type = type;

    public eSlicerSourceType Type { get; }
}

internal class ExcelXmlSource
{
    internal ExcelXmlSource(ZipPackagePart relPart, string relId)
    {
        if (string.IsNullOrEmpty(relId))
        {
            this.Uri = XmlHelper.GetNewUri(relPart.Package, "/xl/slicers/slicer{0}.xml");
            this.Part = relPart.Package.CreatePart(this.Uri, "application/vnd.ms-excel.slicer+xml", CompressionLevel.Default);
            this.Rel = relPart.CreateRelationship(UriHelper.GetRelativeUri(relPart.Uri, this.Uri), TargetMode.Internal, ExcelPackage.schemaRelationshipsSlicer);
            XmlDocument? xml = new XmlDocument();

            XmlHelper.LoadXmlSafe(xml,
                                  "<slicers xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" mc:Ignorable=\"x xr10\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" />",
                                  Encoding.UTF8);

            this.XmlDocument = xml;
        }
        else
        {
            this.Rel = relPart.GetRelationship(relId);
            this.Uri = UriHelper.ResolvePartUri(relPart.Uri, this.Rel.TargetUri);
            this.Part = relPart.Package.GetPart(this.Uri);

            XmlDocument? xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, this.Part.GetStream());
            this.XmlDocument = xml;
        }
    }

    internal ZipPackageRelationship Rel { get; }

    internal ZipPackagePart Part { get; }

    internal Uri Uri { get; }

    public XmlDocument XmlDocument { get; }
}