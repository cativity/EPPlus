﻿/*************************************************************************************************
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// A collection of <see cref="ExcelThreadedCommentMention">mentions</see> that occors in a <see cref="ExcelThreadedComment"/>
/// </summary>
public sealed class ExcelThreadedCommentMentionCollection : XmlHelper, IEnumerable<ExcelThreadedCommentMention>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="nameSpaceManager">The Namespacemangager of the package</param>
    /// <param name="topNode">The <see cref="XmlNode"/> representing the parent element of the collection</param>
    internal ExcelThreadedCommentMentionCollection(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        : base(nameSpaceManager, topNode) =>
        this.LoadMentions();

    private readonly List<ExcelThreadedCommentMention> _mentionList = new List<ExcelThreadedCommentMention>();

    private void LoadMentions()
    {
        foreach (object? mentionNode in this.TopNode.ChildNodes)
        {
            this._mentionList.Add(new ExcelThreadedCommentMention(this.NameSpaceManager, (XmlNode)mentionNode));
        }
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<ExcelThreadedCommentMention> GetEnumerator() => this._mentionList.GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => this._mentionList.GetEnumerator();

    /// <summary>
    /// Adds a mention
    /// </summary>
    /// <param name="person">The <see cref="ExcelThreadedCommentPerson"/> to mention</param>
    /// <param name="textPosition">Index of the first character of the mention in the text</param>
    internal void AddMention(ExcelThreadedCommentPerson person, int textPosition)
    {
        XmlElement? elem = this.TopNode.OwnerDocument.CreateElement("mention", ExcelPackage.schemaThreadedComments);
        _ = this.TopNode.AppendChild(elem);
        ExcelThreadedCommentMention? mention = new ExcelThreadedCommentMention(this.NameSpaceManager, elem);
        mention.MentionId = ExcelThreadedCommentMention.NewId();
        mention.StartIndex = textPosition;

        // + 1 to include the @ prefix...
        mention.Length = person.DisplayName.Length + 1;
        mention.MentionPersonId = person.Id;
        this._mentionList.Add(mention);
    }

    /// <summary>
    /// Rebuilds the collection with the elements sorted by the property StartIndex.
    /// </summary>
    internal void SortAndAddMentionsToXml()
    {
        this._mentionList.Sort((x, y) => x.StartIndex.CompareTo(y.StartIndex));
        this.TopNode.RemoveAll();
        this._mentionList.ForEach(x => this.TopNode.AppendChild(x.TopNode));
    }

    /// <summary>
    /// Remove all mentions from the collection
    /// </summary>
    internal void Clear()
    {
        this._mentionList.Clear();
        this.TopNode.RemoveAll();
    }

    /// <summary>
    ///     Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString() => "Count = " + this._mentionList.Count;
}