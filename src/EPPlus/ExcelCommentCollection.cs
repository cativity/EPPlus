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
using System.Xml;
using System.Collections;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Core.CellStore;
using System.Linq;
using OfficeOpenXml.Core;
using System.Threading;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml;

/// <summary>
/// Collection of Excel Comment objects
/// </summary>  
public class ExcelCommentCollection : IEnumerable, IDisposable
{
    //internal RangeCollection _comments;
    internal List<ExcelComment> _list = new List<ExcelComment>();
    List<int> _listIndex = new List<int>();

    internal ExcelCommentCollection(ExcelPackage pck, ExcelWorksheet ws, XmlNamespaceManager ns)
    {
        this.CommentXml = new XmlDocument();
        this.CommentXml.PreserveWhitespace = false;
        this.NameSpaceManager = ns;
        this.Worksheet = ws;
        this.CreateXml(pck);
        this.AddCommentsFromXml();
    }

    private void CreateXml(ExcelPackage pck)
    {
        ZipPackageRelationshipCollection? commentRels = this.Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaComment);
        bool isLoaded = false;
        this.CommentXml = new XmlDocument();

        foreach (ZipPackageRelationship? commentPart in commentRels)
        {
            this.Uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
            this.Part = pck.ZipPackage.GetPart(this.Uri);
            XmlHelper.LoadXmlSafe(this.CommentXml, this.Part.GetStream());
            this.RelId = commentPart.Id;
            isLoaded = true;
        }

        //Create a new document
        if (!isLoaded)
        {
            this.CommentXml.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors /><commentList /></comments>");
            this.Uri = null;
        }
    }

    private void AddCommentsFromXml()
    {
        //var lst = new List<IRangeID>();
        foreach (XmlElement node in this.CommentXml.SelectNodes("//d:commentList/d:comment", this.NameSpaceManager))
        {
            ExcelComment? comment = new ExcelComment(this.NameSpaceManager, node, new ExcelRangeBase(this.Worksheet, node.GetAttribute("ref")));
            this._listIndex.Add(this._list.Count);
            this.Worksheet._commentsStore.SetValue(comment.Range._fromRow, comment.Range._fromCol, this._list.Count);
            this._list.Add(comment);
        }

        //_comments = new RangeCollection(lst);
    }

    /// <summary>
    /// Access to the comment xml document
    /// </summary>
    public XmlDocument CommentXml { get; set; }

    internal Uri Uri { get; set; }

    internal string RelId { get; set; }

    internal XmlNamespaceManager NameSpaceManager { get; set; }

    internal ZipPackagePart Part { get; set; }

    /// <summary>
    /// A reference to the worksheet object
    /// </summary>
    public ExcelWorksheet Worksheet { get; set; }

    /// <summary>
    /// Number of comments in the collection
    /// </summary>
    public int Count => this._listIndex.Count;

    /// <summary>
    /// Indexer for the comments collection
    /// </summary>
    /// <param name="Index">The index</param>
    /// <returns>The comment</returns>
    public ExcelComment this[int Index]
    {
        get
        {
            if (Index < 0 || Index >= this._listIndex.Count)
            {
                throw new ArgumentOutOfRangeException("Comment index out of range");
            }

            return this._list[this._listIndex[Index]];
        }
    }

    /// <summary>
    /// Indexer for the comments collection
    /// </summary>
    /// <param name="cell">The cell</param>
    /// <returns>The comment</returns>
    public ExcelComment this[ExcelCellAddress cell]
    {
        get
        {
            int i = -1;

            if (this.Worksheet._commentsStore.Exists(cell.Row, cell.Column, ref i))
            {
                return this._list[i];
            }
            else
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Indexer for the comments collection
    /// </summary>
    /// <param name="cellAddress">The cell address</param>
    /// <returns>The comment</returns>
    public ExcelComment this[string cellAddress] => this[new ExcelCellAddress(cellAddress)];

    /// <summary>
    /// Adds a comment to the top left cell of the range
    /// </summary>
    /// <param name="cell">The cell</param>
    /// <param name="Text">The comment text</param>
    /// <param name="author">The author for the comment. If this property is null or blank EPPlus will set it to the identity of the ClaimsPrincipal if available otherwise to "Anonymous"</param>
    /// <returns>The comment</returns>
    public ExcelComment Add(ExcelRangeBase cell, string Text, string author = null)
    {
        if (string.IsNullOrEmpty(author))
        {
#if Core
            author = System.Security.Claims.ClaimsPrincipal.Current?.Identity?.Name;
#else
                author = Thread.CurrentPrincipal?.Identity?.Name;
#endif
            if (string.IsNullOrEmpty(author))
            {
                author = "Anonymous";
            }
        }

        XmlElement? elem = this.CommentXml.CreateElement("comment", ExcelPackage.schemaMain);

        //int ix=_comments.IndexOf(ExcelAddress.GetCellID(Worksheet.SheetID, cell._fromRow, cell._fromCol));
        //Make sure the nodes come on order.
        int row = cell.Start.Row,
            column = cell.Start.Column;

        ExcelComment nextComment = null;

        if (this.Worksheet._commentsStore.NextCell(ref row, ref column))
        {
            nextComment = this._list[this.Worksheet._commentsStore.GetValue(row, column)];
        }

        if (nextComment == null)
        {
            this.CommentXml.SelectSingleNode("d:comments/d:commentList", this.NameSpaceManager).AppendChild(elem);
        }
        else
        {
            nextComment._commentHelper.TopNode.ParentNode.InsertBefore(elem, nextComment._commentHelper.TopNode);
        }

        elem.SetAttribute("ref", cell.Start.Address);
        ExcelComment comment = new ExcelComment(this.NameSpaceManager, elem, cell);
        comment.RichText.Add(Text);
        comment.Author = author;
        this._listIndex.Add(this._list.Count);
        this.Worksheet._commentsStore.SetValue(cell.Start.Row, cell.Start.Column, this._list.Count);
        this._list.Add(comment);

        //Check if a value exists otherwise add one so it is saved when the cells collection is iterated
        if (!this.Worksheet.ExistsValueInner(cell._fromRow, cell._fromCol))
        {
            this.Worksheet.SetValueInner(cell._fromRow, cell._fromCol, null);
        }

        return comment;
    }

    /// <summary>
    /// Removes the comment
    /// </summary>
    /// <param name="comment">The comment to remove</param>
    public void Remove(ExcelComment comment) => this.Remove(comment, false);

    internal void Remove(ExcelComment comment, bool shift)
    {
        int i = -1;
        ExcelComment c = null;

        if (this.Worksheet._commentsStore.Exists(comment.Range._fromRow, comment.Range._fromCol, ref i))
        {
            c = this._list[i];
        }

        if (comment == c)
        {
            _ = comment.TopNode.ParentNode.RemoveChild(comment.TopNode); //Remove VML
            _ = comment._commentHelper.TopNode.ParentNode.RemoveChild(comment._commentHelper.TopNode); //Remove Comment

            this.Worksheet.VmlDrawings._drawingsCellStore.Delete(comment.Range._fromRow, comment.Range._fromCol, 1, 1, shift);
            this.Worksheet._commentsStore.Delete(comment.Range._fromRow, comment.Range._fromCol, 1, 1, shift);
            this._list[i] = null;
            _ = this._listIndex.Remove(i);

            //if(_listIndex.Count==0)
            //{
            //    _list.Clear();
            //}
        }
        else
        {
            throw new ArgumentException("Comment does not exist in the worksheet");
        }
    }

    /// <summary>
    /// Shifts all comments based on their address and the location of inserted rows and columns.
    /// </summary>
    /// <param name="fromRow">The start row.</param>
    /// <param name="fromCol">The start column.</param>
    /// <param name="rows">The number of rows to insert.</param>
    /// <param name="columns">The number of columns to insert.</param>
    /// <param name="toRow">If the delete is in a range, this is the end row</param>
    /// <param name="toCol">If the delete is in a range, this the end column</param>
    internal void Delete(int fromRow, int fromCol, int rows, int columns, int toRow = ExcelPackage.MaxRows, int toCol = ExcelPackage.MaxColumns)
    {
        List<ExcelComment> deletedComments = new List<ExcelComment>();

        foreach (ExcelComment comment in this._list.Where(x => x != null))
        {
            ExcelAddressBase address = new(comment.Address);

            if (columns > 0 && address._fromCol >= fromCol && address._fromRow >= fromRow && address._toRow <= toRow)
            {
                address = address.DeleteColumn(fromCol, columns);
            }

            if (rows > 0 && address._fromRow >= fromRow && address._fromCol >= fromCol && address._toCol <= toCol)
            {
                address = address.DeleteRow(fromRow, rows);
            }

            if (address == null || address.Address == "#REF!")
            {
                deletedComments.Add(comment);
            }
            else
            {
                comment.Reference = address.Address;
            }
        }

        foreach (ExcelComment? comment in deletedComments)
        {
            _ = comment.TopNode.ParentNode.RemoveChild(comment.TopNode); //Remove VML
            _ = comment._commentHelper.TopNode.ParentNode.RemoveChild(comment._commentHelper.TopNode); //Remove Comment
            int ix = this._list.IndexOf(comment);
            this._list[ix] = null;
            _ = this._listIndex.Remove(ix);
        }
    }

    /// <summary>
    /// Shifts all comments based on their address and the location of inserted rows and columns.
    /// </summary>
    /// <param name="fromRow">The start row</param>
    /// <param name="fromCol">The start column</param>
    /// <param name="rows">The number of rows to insert</param>
    /// <param name="columns">The number of columns to insert</param>
    /// <param name="toRow">If the insert is in a range, this is the end row</param>
    /// <param name="toCol">If the insert is in a range, this the end column</param>
    internal void Insert(int fromRow, int fromCol, int rows, int columns, int toRow = ExcelPackage.MaxRows, int toCol = ExcelPackage.MaxColumns)
    {
        foreach (ExcelComment comment in this._list.Where(x => x != null))
        {
            ExcelAddressBase? address = new ExcelAddressBase(comment.Address);

            if (rows > 0 && address._fromRow >= fromRow && address._fromCol >= fromCol && address._toCol <= toCol)
            {
                comment.Reference = comment.Range.AddRow(fromRow, rows).Address;
            }

            if (columns > 0 && address._fromCol >= fromCol && address._fromRow >= fromRow && address._toRow <= toRow)
            {
                comment.Reference = comment.Range.AddColumn(fromCol, columns).Address;
            }
        }
    }

    void IDisposable.Dispose()
    {
    }

    /// <summary>
    /// Removes the comment at the specified position
    /// </summary>
    /// <param name="Index">The index</param>
    public void RemoveAt(int Index) => this.Remove(this[Index]);

    #region IEnumerable Members

    IEnumerator IEnumerable.GetEnumerator() => this._list.Where(x => x != null).GetEnumerator();

    #endregion

    internal void Clear()
    {
        while (this.Count > 0)
        {
            this.RemoveAt(0);
        }
    }
}