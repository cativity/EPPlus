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
using System.Text;
using System.Xml;
using System.Collections;
using System.Globalization;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// A collection of vml drawings used for header and footer picturess
/// </summary>
public class ExcelVmlDrawingPictureCollection : ExcelVmlDrawingBaseCollection, IEnumerable
{
    internal List<ExcelVmlDrawingPicture> _images;

    internal ExcelVmlDrawingPictureCollection(ExcelWorksheet ws, Uri uri)
        : base(ws, uri, "d:legacyDrawingHF/@r:id")
    {
        if (uri == null)
        {
            this.VmlDrawingXml.LoadXml(CreateVmlDrawings());
            this._images = new List<ExcelVmlDrawingPicture>();
        }
        else
        {
            this.AddDrawingsFromXml();
        }
    }

    private void AddDrawingsFromXml()
    {
        XmlNodeList? nodes = this.VmlDrawingXml.SelectNodes("//v:shape", this.NameSpaceManager);
        this._images = new List<ExcelVmlDrawingPicture>();

        foreach (XmlNode node in nodes)
        {
            ExcelVmlDrawingPicture? img = new ExcelVmlDrawingPicture(node, this.NameSpaceManager, this._ws);
            ZipPackageRelationship? rel = this.Part.GetRelationship(img.RelId);
            img.ImageUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            this._images.Add(img);
        }
    }

    private static string CreateVmlDrawings()
    {
        string vml = string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">",
                                   ExcelPackage.schemaMicrosoftVml,
                                   ExcelPackage.schemaMicrosoftOffice,
                                   ExcelPackage.schemaMicrosoftExcel);

        vml += "<o:shapelayout v:ext=\"edit\">";
        vml += "<o:idmap v:ext=\"edit\" data=\"1\"/>";
        vml += "</o:shapelayout>";

        vml += "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
        vml += "<v:stroke joinstyle=\"miter\" />";
        vml += "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
        vml += "</v:shapetype>";
        vml += "</xml>";

        return vml;
    }

    internal ExcelVmlDrawingPicture Add(string id, Uri uri, string name, double width, double height)
    {
        XmlNode node = this.AddImage(id, uri, name, width, height);
        ExcelVmlDrawingPicture? draw = new ExcelVmlDrawingPicture(node, this.NameSpaceManager, this._ws);
        draw.ImageUri = uri;
        this._images.Add(draw);

        return draw;
    }

    private XmlNode AddImage(string id, Uri targeUri, string Name, double width, double height)
    {
        XmlElement? node = this.VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);
        _ = this.VmlDrawingXml.DocumentElement.AppendChild(node);
        node.SetAttribute("id", id);
        node.SetAttribute("o:type", "#_x0000_t75");

        node.SetAttribute("style",
                          string.Format("position:absolute;margin-left:0;margin-top:0;width:{0}pt;height:{1}pt;z-index:1",
                                        width.ToString(CultureInfo.InvariantCulture),
                                        height.ToString(CultureInfo.InvariantCulture)));

        node.InnerXml = string.Format("<v:imagedata o:relid=\"\" o:title=\"{0}\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>", Name);

        return node;
    }

    /// <summary>
    /// Indexer
    /// </summary>
    /// <param name="Index">Index</param>
    /// <returns>The VML Drawing Picture object</returns>
    public ExcelVmlDrawingPicture this[int Index] => this._images[Index] as ExcelVmlDrawingPicture;

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._images.Count;

    int _nextID;

    /// <summary>
    /// returns the next drawing id.
    /// </summary>
    /// <returns></returns>
    internal string GetNewId()
    {
        if (this._nextID == 0)
        {
            foreach (ExcelVmlDrawingComment draw in this)
            {
                if (draw.Id.Length > 3 && draw.Id.StartsWith("vml", StringComparison.OrdinalIgnoreCase))
                {
                    if (int.TryParse(draw.Id.Substring(3, draw.Id.Length - 3), NumberStyles.Number, CultureInfo.InvariantCulture, out int id))
                    {
                        if (id > this._nextID)
                        {
                            this._nextID = id;
                        }
                    }
                }
            }
        }

        this._nextID++;

        return "vml" + this._nextID.ToString();
    }

    #region IEnumerable Members

    IEnumerator IEnumerable.GetEnumerator() => this._images.GetEnumerator();

    #endregion
}