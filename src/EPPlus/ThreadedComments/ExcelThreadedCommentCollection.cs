/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/29/2020         EPPlus Software AB       Threaded comments
 *************************************************************************************************/

using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// This class represents an enumerable of <see cref="ExcelThreadedComment"/>s.
/// </summary>
public class ExcelThreadedCommentCollection : XmlHelper, IEnumerable<ExcelThreadedComment>
{
    internal ExcelThreadedCommentCollection(ExcelWorksheet worksheet, XmlNode topNode)
        : base(worksheet.NameSpaceManager, topNode)
    {
        //this._package = worksheet._package;
        this.Worksheet = worksheet;
    }

    //private readonly ExcelPackage _package;

    internal Uri Uri { get; set; }

    internal string RelId { get; set; }

    internal Packaging.ZipPackagePart Part { get; set; }

    private readonly Dictionary<string, ExcelThreadedComment> _commentsIndex = new Dictionary<string, ExcelThreadedComment>();
    private readonly List<ExcelThreadedComment> _commentList = new List<ExcelThreadedComment>();

    /// <summary>
    /// A reference to the worksheet object
    /// </summary>
    public ExcelWorksheet Worksheet { get; set; }

    /// <summary>
    /// Returns a <see cref="ExcelThreadedComment"/> by its index
    /// </summary>
    /// <param name="index">Index in this collection</param>
    /// <returns>The <see cref="ExcelThreadedComment"/> at the requested <paramref name="index"/></returns>
    /// <exception cref="ArgumentOutOfRangeException">If the <paramref name="index"/> falls out of range</exception>
    public ExcelThreadedComment this[int index]
    {
        get { return this._commentList[index]; }
    }

    /// <summary>
    /// Returns a <see cref="ExcelThreadedComment"/> by its <paramref name="id"/>
    /// </summary>
    /// <param name="id">Id of the requested <see cref="ExcelThreadedComment"/></param>
    /// <returns>The requested <see cref="ExcelThreadedComment"/></returns>
    /// <exception cref="ArgumentException">If the requested <paramref name="id"/> was not present.</exception>
    public ExcelThreadedComment this[string id]
    {
        get
        {
            if (!this._commentsIndex.ContainsKey(id))
            {
                throw new ArgumentException("Id " + id + " was not present in the comments.");
            }

            return this._commentsIndex[id];
        }
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<ExcelThreadedComment> GetEnumerator()
    {
        return this._commentList.GetEnumerator();
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._commentList.GetEnumerator();
    }

    /// <summary>
    /// Number of <see cref="ExcelThreadedComment"/>s
    /// </summary>
    public int Count
    {
        get { return this._commentList.Count; }
    }

    private void RebuildIndex()
    {
        this._commentsIndex.Clear();

        foreach (ExcelThreadedComment? comment in this._commentList)
        {
            this._commentsIndex[comment.Id] = comment;
        }
    }

    internal void Add(ExcelThreadedComment comment)
    {
        this._commentList.Add(comment);

        if (this.TopNode.SelectSingleNode("tc:threadedComment[@id='" + comment.Id + "']", this.NameSpaceManager) == null)
        {
            _ = this.TopNode.AppendChild(comment.TopNode);
        }

        this.RebuildIndex();
    }

    internal bool Remove(ExcelThreadedComment comment)
    {
        int index = this._commentList.IndexOf(comment);
        _ = this._commentList.Remove(comment);
        XmlNode? commentNode = this.TopNode.SelectSingleNode("tc:threadedComment[@id='" + comment.Id + "']", this.NameSpaceManager);

        if (commentNode != null)
        {
            _ = this.TopNode.RemoveChild(commentNode);

            //Reset the parentid to the first item in the list if we remove the first comment
            if (index == 0 && this._commentList.Count > 0)
            {
                ((XmlElement)this._commentList[0].TopNode).RemoveAttribute("parentId");

                for (int i = 1; i < this._commentList.Count; i++)
                {
                    this._commentList[i].ParentId = this._commentList[0].Id;
                }
            }

            this.RebuildIndex();

            return true;
        }

        return false;
    }

    /// <summary>
    /// Removes all <see cref="ExcelThreadedComment"/>s in the collection
    /// </summary>
    internal void Clear()
    {
        foreach (XmlNode? node in this._commentList.Select(x => x.TopNode))
        {
            _ = this.TopNode.RemoveChild(node);
        }

        this._commentList.Clear();
    }

    /// <summary>
    ///     Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString()
    {
        return "Count = " + this._commentList.Count;
    }
}