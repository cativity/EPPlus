/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/

using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.ExternalReferences;

/// <summary>
/// A collection of external links referenced by the workbook.
/// </summary>
public class ExcelExternalLinksCollection : IEnumerable<ExcelExternalLink>
{
    readonly List<ExcelExternalLink> _list = new List<ExcelExternalLink>();
    readonly ExcelWorkbook _wb;

    internal ExcelExternalLinksCollection(ExcelWorkbook wb)
    {
        this._wb = wb;
        this.LoadExternalReferences();
    }

    internal void AddInternal(ExcelExternalLink externalLink) => this._list.Add(externalLink);

    /// <summary>
    ///     Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<ExcelExternalLink> GetEnumerator() => this._list.GetEnumerator();

    /// <summary>
    ///     Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => this._list.GetEnumerator();

    /// <summary>
    /// Gets the number of items in the collection
    /// </summary>
    public int Count => this._list.Count;

    /// <summary>
    /// The indexer for the collection
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns></returns>
    public ExcelExternalLink this[int index] => this._list[index];

    /// <summary>
    /// Adds an external reference to another workbook. 
    /// </summary>
    /// <param name="file">The location of the external workbook. The external workbook must of type .xlsx, .xlsm or xlst</param>
    /// <returns>The <see cref="ExcelExternalWorkbook"/> object</returns>
    public ExcelExternalWorkbook AddExternalWorkbook(FileInfo file)
    {
        if (file == null || file.Exists == false)
        {
            throw new FileNotFoundException("The file does not exist.");
        }

        ExcelPackage? p = new ExcelPackage(file);
        ExcelExternalWorkbook? ewb = new ExcelExternalWorkbook(this._wb, p);
        this._list.Add(ewb);

        return ewb;
    }

    internal void LoadExternalReferences()
    {
        XmlNodeList nl = this._wb.WorkbookXml.SelectNodes("//d:externalReferences/d:externalReference", this._wb.NameSpaceManager);

        if (nl != null)
        {
            foreach (XmlElement elem in nl)
            {
                string rID = elem.GetAttribute("r:id");
                ZipPackageRelationship? rel = this._wb.Part.GetRelationship(rID);
                ZipPackagePart? part = this._wb._package.ZipPackage.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                XmlTextReader? xr = new XmlTextReader(part.GetStream());

                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        switch (xr.Name)
                        {
                            case "externalBook":
                                this.AddInternal(new ExcelExternalWorkbook(this._wb, xr, part, elem));

                                break;

                            case "ddeLink":
                                this.AddInternal(new ExcelExternalDdeLink(this._wb, xr, part, elem));

                                break;

                            case "oleLink":
                                this.AddInternal(new ExcelExternalOleLink(this._wb, xr, part, elem));

                                break;

                            case "extLst":

                                break;

                            default:
                                break;
                        }
                    }
                }

                xr.Close();
            }
        }
    }

    /// <summary>
    /// Removes the external link at the zero-based index. If the external reference is an workbook any formula links are broken.
    /// </summary>
    /// <param name="index">The zero-based index</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= this._list.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        this.Remove(this._list[index]);
    }

    /// <summary>
    /// Removes the external link from the package.If the external reference is an workbook any formula links are broken.
    /// </summary>
    /// <param name="externalLink"></param>
    public void Remove(ExcelExternalLink externalLink)
    {
        int ix = this._list.IndexOf(externalLink);

        this._wb._package.ZipPackage.DeletePart(externalLink.Part.Uri);

        if (externalLink.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
        {
            ExternalLinksHandler.BreakFormulaLinks(this._wb, ix, true);
        }

        XmlNode? extRefs = externalLink.WorkbookElement.ParentNode;
        _ = (extRefs?.RemoveChild(externalLink.WorkbookElement));

        if (extRefs?.ChildNodes.Count == 0)
        {
            _ = (extRefs.ParentNode?.RemoveChild(extRefs));
        }

        _ = this._list.Remove(externalLink);
    }

    /// <summary>
    /// Clear all external links and break any formula links.
    /// </summary>
    public void Clear()
    {
        if (this._list.Count == 0)
        {
            return;
        }

        XmlNode? extRefs = this._list[0].WorkbookElement.ParentNode;

        ExternalLinksHandler.BreakAllFormulaLinks(this._wb);

        while (this._list.Count > 0)
        {
            this._wb._package.ZipPackage.DeletePart(this._list[0].Part.Uri);
            this._list.RemoveAt(0);
        }

        _ = (extRefs?.ParentNode?.RemoveChild(extRefs));
    }

    /// <summary>
    /// A list of directories to look for the external files that cannot be found on the path of the uri.
    /// </summary>
    public List<DirectoryInfo> Directories { get; } = new List<DirectoryInfo>();

    /// <summary>
    /// Will load all external workbooks that can be accessed via the file system.
    /// External workbook referenced via other protocols must be loaded manually.
    /// </summary>
    /// <returns>Returns false if any workbook fails to loaded otherwise true. </returns>
    public bool LoadWorkbooks()
    {
        bool ret = true;

        foreach (ExcelExternalLink? link in this._list)
        {
            if (link.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
            {
                ExcelExternalWorkbook? externalWb = link.As.ExternalWorkbook;

                if (externalWb.Package == null)
                {
                    if (externalWb.Load() == false)
                    {
                        ret = false;
                    }
                }
            }
        }

        return ret;
    }

    internal int GetExternalLink(string extRef)
    {
        if (string.IsNullOrEmpty(extRef))
        {
            return -1;
        }

        if (extRef.Any(c => char.IsDigit(c) == false))
        {
            if (ExcelExternalLink.HasWebProtocol(extRef))
            {
                for (int ix = 0; ix < this._list.Count; ix++)
                {
                    if (this._list[ix].ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                    {
                        if (extRef.Equals(this._list[ix].As.ExternalWorkbook.ExternalLinkUri.OriginalString, StringComparison.OrdinalIgnoreCase))
                        {
                            return ix;
                        }
                    }
                }

                return -1;
            }

            if (extRef.StartsWith("file:///"))
            {
                extRef = extRef.Substring(8);
            }

            int ret = -1;

            try
            {
                FileInfo? fi = new FileInfo(extRef);

                for (int ix = 0; ix < this._list.Count; ix++)
                {
                    if (this._list[ix].ExternalLinkType == eExternalLinkType.ExternalWorkbook)
                    {
                        ExcelExternalWorkbook? wb = this._list[ix].As.ExternalWorkbook;

                        if (wb.File == null)
                        {
                            string? fileName = wb.ExternalLinkUri?.OriginalString;

                            if (ExcelExternalLink.HasWebProtocol(fileName))
                            {
                                if (fileName.Equals(extRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    return ix;
                                }

                                continue;
                            }
                        }

                        if (fi.Name.Equals(wb.File.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            ret = ix;
                        }
                    }
                }
            }
            catch //If the FileInfo is 
            {
                return -1;
            }

            return ret;
        }
        else
        {
            int ix = int.Parse(extRef) - 1;

            if (ix < this._list.Count)
            {
                return ix;
            }
        }

        return -1;
    }

    internal int GetIndex(ExcelExternalLink link) => this._list.IndexOf(link);

    /// <summary>
    /// Updates the value cache for any external workbook in the collection. The link must be an workbook and of type xlsx, xlsm or xlst.
    /// </summary>
    /// <returns>True if all updates succeeded, otherwise false. Any errors can be found on the External links. <seealso cref="ExcelExternalLink.ErrorLog"/></returns>
    public bool UpdateCaches()
    {
        bool ret = true;

        foreach (ExcelExternalLink? er in this._list)
        {
            if (er.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
            {
                if (er.As.ExternalWorkbook.UpdateCache() == false)
                {
                    ret = false;
                }
            }
        }

        return ret;
    }
}