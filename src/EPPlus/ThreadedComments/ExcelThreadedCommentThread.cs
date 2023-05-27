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

using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// Represents a thread of <see cref="ExcelThreadedComment"/>s in a cell on a worksheet. Contains functionality to add and modify these comments.
/// </summary>
public class ExcelThreadedCommentThread
{
    internal ExcelThreadedCommentThread(ExcelCellAddress cellAddress, XmlDocument commentsXml, ExcelWorksheet worksheet)
    {
        this.CellAddress = cellAddress;
        this.ThreadedCommentsXml = commentsXml;
        this.Worksheet = worksheet;
        this.Comments = new ExcelThreadedCommentCollection(worksheet, commentsXml.SelectSingleNode("tc:ThreadedComments", worksheet.NameSpaceManager));
    }

    /// <summary>
    /// The address of the cell of the comment thread
    /// </summary>
    public ExcelCellAddress CellAddress { get; internal set; }

    /// <summary>
    /// A collection of comments in the thread.
    /// </summary>
    public ExcelThreadedCommentCollection Comments { get; private set; }

    /// <summary>
    /// The worksheet where this comment thread resides
    /// </summary>
    public ExcelWorksheet Worksheet { get; private set; }

    /// <summary>
    /// The raw xml representing this comment thread.
    /// </summary>
    public XmlDocument ThreadedCommentsXml { get; private set; }

    private void ReplicateThreadToLegacyComment()
    {
        IEnumerable<ExcelThreadedComment>? tc = this.Comments as IEnumerable<ExcelThreadedComment>;

        if (!tc.Any())
        {
            return;
        }

        int tcIndex = 0;
        StringBuilder? commentText = new StringBuilder();
        string? authorId = "tc=" + tc.First().Id;
        commentText.AppendLine("This comment reflects a threaded comment in this cell, a feature that might be supported by newer versions of your spreadsheet program (for example later versions of Excel). Any edits will be overwritten if opened in a spreadsheet program that supports threaded comments.");
        commentText.AppendLine();

        foreach (ExcelThreadedComment? threadedComment in tc)
        {
            if (tcIndex == 0)
            {
                commentText.AppendLine("Comment:");
            }
            else
            {
                commentText.AppendLine("Reply:");
            }

            commentText.AppendLine(threadedComment.Text);
            tcIndex++;
        }

        ExcelComment? comment = this.Worksheet.Comments[this.CellAddress];

        if (comment == null)
        {
            this.Worksheet.Comments.Add(this.Worksheet.Cells[this.CellAddress.Address], commentText.ToString(), authorId);
        }
        else
        {
            comment.Text = commentText.ToString();
            comment.Author = authorId;
        }
    }

    /// <summary>
    /// When this method is called the legacy comment representing the thread will be rebuilt.
    /// </summary>
    internal void OnCommentThreadChanged()
    {
        this.ReplicateThreadToLegacyComment();
    }

    /// <summary>
    /// Adds a <see cref="ExcelThreadedComment"/> to the thread
    /// </summary>
    /// <param name="personId">Id of the author, see <see cref="ExcelThreadedCommentPerson"/></param>
    /// <param name="text">Text of the comment</param>
    public ExcelThreadedComment AddComment(string personId, string text)
    {
        return this.AddComment(personId, text, true);
    }

    internal ExcelThreadedComment AddComment(string personId, string text, bool replicateLegacyComment)
    {
        Require.That(text).Named("text").IsNotNullOrEmpty();
        Require.That(personId).Named("personId").IsNotNullOrEmpty();
        string? parentId = string.Empty;

        if (this.Comments.Any())
        {
            parentId = this.Comments.First().Id;
        }

        XmlElement? xmlNode = this.ThreadedCommentsXml.CreateElement("threadedComment", ExcelPackage.schemaThreadedComments);
        this.ThreadedCommentsXml.SelectSingleNode("tc:ThreadedComments", this.Worksheet.NameSpaceManager).AppendChild(xmlNode);
        ExcelThreadedComment? newComment = new ExcelThreadedComment(xmlNode, this.Worksheet.NameSpaceManager, this.Worksheet.Workbook, this);
        newComment.Id = ExcelThreadedComment.NewId();
        newComment.CellAddress = new ExcelCellAddress(this.CellAddress.Address);
        newComment.Text = text;
        newComment.PersonId = personId;
        newComment.DateCreated = DateTime.Now;

        if (!string.IsNullOrEmpty(parentId))
        {
            newComment.ParentId = parentId;
        }

        this.Comments.Add(newComment);

        if (replicateLegacyComment)
        {
            this.ReplicateThreadToLegacyComment();
        }

        return newComment;
    }

    internal void AddComment(ExcelThreadedComment comment)
    {
        this.Comments.Add(comment);
        this.ReplicateThreadToLegacyComment();
    }

    /// <summary>
    /// Adds a <see cref="ExcelThreadedComment"/> with mentions in the text to the thread.
    /// </summary>
    /// <param name="personId">Id of the <see cref="ExcelThreadedCommentPerson">autor</see></param>
    /// <param name="textWithFormats">A string with format placeholders - same as in string.Format. Index in these should correspond to an index in the <paramref name="personsToMention"/> array.</param>
    /// <param name="personsToMention">A params array of <see cref="ExcelThreadedCommentPerson"/>. Their DisplayName property will be used to replace the format placeholders.</param>
    /// <returns>The added <see cref="ExcelThreadedComment"/></returns>
    public ExcelThreadedComment AddComment(string personId, string textWithFormats, params ExcelThreadedCommentPerson[] personsToMention)
    {
        ExcelThreadedComment? comment = this.AddComment(personId, textWithFormats, true);
        MentionsHelper.InsertMentions(comment, textWithFormats, personsToMention);

        return comment;
    }

    /// <summary>
    /// Removes a <see cref="ExcelThreadedComment"/> from the thread.
    /// </summary>
    /// <param name="comment">The comment to remove</param>
    /// <returns>true if the comment was removed, otherwise false</returns>
    public bool Remove(ExcelThreadedComment comment)
    {
        if (this.Comments.Remove(comment))
        {
            this.ReplicateThreadToLegacyComment();

            return true;
        }

        return false;
    }

    /// <summary>
    /// Closes the thread, only the author can re-open it.
    /// </summary>
    public void ResolveThread()
    {
        if (!this.Comments.Any())
        {
            throw new InvalidOperationException("Cannot resolve an empty thread (it has not comments)");
        }

        this.Comments.First().Done = true;
    }

    /// <summary>
    /// If true the thread is resolved, i.e. closed for edits or further comments.
    /// </summary>
    public bool IsResolved
    {
        get
        {
            if (!this.Comments.Any())
            {
                return false;
            }

            return this.Comments.First().Done ?? false;
        }
    }

    /// <summary>
    /// Re-opens a resolved thread.
    /// </summary>
    public void ReopenThread()
    {
        if (!this.Comments.Any())
        {
            throw new InvalidOperationException("Cannot re-open an empty thread (it has not comments)");
        }

        this.Comments.First().Done = null;
    }

    /// <summary>
    /// Deletes all <see cref="ExcelThreadedComment"/>s in the thread and the legacy <see cref="ExcelComment"/> in the cell.
    /// </summary>
    public void DeleteThread()
    {
        this.Comments.Clear();

        if (this.Worksheet.Comments[this.CellAddress] != null)
        {
            ExcelComment? comment = this.Worksheet.Comments[this.CellAddress];
            this.Worksheet.Comments.Remove(comment);
        }
    }

    internal void AddCommentFromXml(XmlElement copyFromElement)
    {
        XmlElement? xmlNode = this.ThreadedCommentsXml.CreateElement("threadedComment", ExcelPackage.schemaThreadedComments);
        this.ThreadedCommentsXml.SelectSingleNode("tc:ThreadedComments", this.Worksheet.NameSpaceManager).AppendChild(xmlNode);

        foreach (XmlAttribute attr in copyFromElement.Attributes)
        {
            if (attr.LocalName == "ref")
            {
                xmlNode.SetAttribute("ref", this.CellAddress.Address);
            }
            else if (attr.LocalName == "id")
            {
                xmlNode.SetAttribute("id", ExcelThreadedComment.NewId());
            }
            else
            {
                xmlNode.SetAttribute(attr.LocalName, attr.Value);
            }
        }

        xmlNode.InnerXml = copyFromElement.InnerXml;
        ExcelThreadedComment? tc = new ExcelThreadedComment(xmlNode, this.Worksheet.NameSpaceManager, this.Worksheet.Workbook, this);

        if (this.Comments.Count > 0)
        {
            tc.ParentId = this.Comments[0].Id;
        }

        foreach (ExcelThreadedCommentMention? m in tc.Mentions)
        {
            m.MentionId = ExcelThreadedComment.NewId();
        }

        this.AddComment(tc);
    }

    /// <summary>
    ///     Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString()
    {
        return "Count = " + this.Comments.Count;
    }
}