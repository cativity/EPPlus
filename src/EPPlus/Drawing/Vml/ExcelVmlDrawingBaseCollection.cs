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
using OfficeOpenXml.Utils;
using System.IO;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Base collection for VML drawings
/// </summary>
public class ExcelVmlDrawingBaseCollection
{        
    internal ExcelPackage _package;
    internal ExcelWorksheet _ws;
    internal ExcelVmlDrawingBaseCollection(ExcelWorksheet ws, Uri uri, string worksheetRelIdPath)
    {
        this.VmlDrawingXml = new XmlDocument();
        this.VmlDrawingXml.PreserveWhitespace = false;
            
        NameTable nt=new NameTable();
        this.NameSpaceManager = new XmlNamespaceManager(nt);
        this.NameSpaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
        this.NameSpaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
        this.NameSpaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
        this.Uri = uri;
        this._package = ws.Workbook._package;
        this._ws = ws;
        if (uri == null)
        {
            int id = this._ws.SheetId;
        }
        else
        {
            this.Part= this._package.ZipPackage.GetPart(uri);
            try
            {                    
                XmlHelper.LoadXmlSafe(this.VmlDrawingXml, this.Part.GetStream());
            }
            catch
            {
                //VML can contain unclosed br tags. Try handle this.
                string? xml = new StreamReader(this.Part.GetStream()).ReadToEnd();
                XmlHelper.LoadXmlSafe(this.VmlDrawingXml, RemoveUnclosedBrTags(xml), Encoding.UTF8);
            }
        }
    }

    private static string RemoveUnclosedBrTags(string xml)
    {
        //TODO:Vml can contain unclosed BR tags. Replace with correctly closed tag and retry. Replace this code with a better approach.
        return xml.Replace("</br>", "").Replace("<br>", "<br/>");
    }

    internal XmlDocument VmlDrawingXml { get; set; }
    internal Uri Uri { get; set; }
    internal string RelId { get; set; }
    internal ZipPackagePart Part { get; set; }
    internal XmlNamespaceManager NameSpaceManager { get; set; }
    internal void CreateVmlPart()
    {
        if (this.Uri == null)
        {
            int id = this._ws.SheetId;
            this.Uri = XmlHelper.GetNewUri(this._package.ZipPackage, @"/xl/drawings/vmlDrawing{0}.vml", ref id);
        }
        if (this.Part == null)
        {
            this.Part = this._package.ZipPackage.CreatePart(this.Uri, ContentTypes.contentTypeVml, this._package.Compression);
            ZipPackageRelationship? rel = this._ws.Part.CreateRelationship(UriHelper.GetRelativeUri(this._ws.WorksheetUri, this.Uri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
            this._ws.SetXmlNodeString("d:legacyDrawing/@r:id", rel.Id);
            this.RelId = rel.Id;
        }

        this.VmlDrawingXml.Save(this.Part.GetStream(FileMode.Create));
    }
}