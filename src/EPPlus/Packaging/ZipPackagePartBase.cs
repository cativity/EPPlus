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
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Packaging.Ionic.Zlib;

namespace OfficeOpenXml.Packaging;

/// <summary>
/// Baseclass for a relation ship between two parts in a package
/// </summary>
public abstract class ZipPackagePartBase
{
    /// <summary>
    /// Relationships collection
    /// </summary>
    protected internal ZipPackageRelationshipCollection _rels = new ZipPackageRelationshipCollection();

    int maxRId = 1;

    internal void DeleteRelationship(string id)
    {
        this._rels.Remove(id);
        UpdateMaxRId(id, ref this.maxRId);
    }

    /// <summary>
    /// Updates the maximum id for the relationship
    /// </summary>
    /// <param name="id">The Id</param>
    /// <param name="maxRId">Return the maximum relation id</param>
    protected internal static void UpdateMaxRId(string id, ref int maxRId)
    {
        if (id.StartsWith("rId"))
        {
            if (int.TryParse(id.Substring(3), out int num))
            {
                if (num == maxRId - 1)
                {
                    maxRId--;
                }
            }
        }
    }

    internal virtual ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
    {
        ZipPackageRelationship? rel = new ZipPackageRelationship();
        rel.TargetUri = targetUri;
        rel.TargetMode = targetMode;
        rel.RelationshipType = relationshipType;
        rel.Id = "rId" + (this.maxRId++).ToString();
        this._rels.Add(rel);

        return rel;
    }

    internal virtual ZipPackageRelationship CreateRelationship(string target, TargetMode targetMode, string relationshipType)
    {
        ZipPackageRelationship? rel = new ZipPackageRelationship();
        rel.Target = target;
        rel.TargetMode = targetMode;
        rel.RelationshipType = relationshipType;
        rel.Id = "rId" + (this.maxRId++).ToString();
        this._rels.Add(rel);

        return rel;
    }

    internal bool RelationshipExists(string id) => this._rels.ContainsKey(id);

    internal ZipPackageRelationshipCollection GetRelationshipsByType(string schema) => this._rels.GetRelationshipsByType(schema);

    internal ZipPackageRelationshipCollection GetRelationships() => this._rels;

    internal ZipPackageRelationship GetRelationship(string id) => this._rels[id];

    internal void ReadRelation(string xml, string source)
    {
        XmlDocument? doc = new XmlDocument();
        XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

        foreach (XmlElement c in doc.DocumentElement.ChildNodes)
        {
            string? target = c.GetAttribute("Target");
            ZipPackageRelationship? rel = new ZipPackageRelationship();
            rel.Id = c.GetAttribute("Id");
            rel.RelationshipType = c.GetAttribute("Type");
            rel.TargetMode = c.GetAttribute("TargetMode").Equals("external", StringComparison.OrdinalIgnoreCase) ? TargetMode.External : TargetMode.Internal;

            if (target.StartsWith("#"))
            {
                rel.Target = c.GetAttribute("Target");
            }
            else
            {
                try
                {
                    rel.TargetUri = new Uri(c.GetAttribute("Target"), UriKind.RelativeOrAbsolute);
                }
                catch
                {
                    //The URI is not a valid URI. Encode it to make i valid.
                    rel.TargetUri = new Uri("Invalid:URI " + Uri.EscapeDataString(c.GetAttribute("Target")), UriKind.RelativeOrAbsolute);
                    rel.Target = c.GetAttribute("Target");
                }
            }

            if (!string.IsNullOrEmpty(source))
            {
                rel.SourceUri = new Uri(source, UriKind.Relative);
            }

            if (rel.Id.StartsWith("rid", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(rel.Id.Substring(3), out int id))
                {
                    if (id >= this.maxRId && id < int.MaxValue - 10000) //Not likly to have this high id's but make sure we have space to avoid overflow.
                    {
                        this.maxRId = id + 1;
                    }
                }
            }

            this._rels.Add(rel);
        }
    }
}