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

using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// Accessor for <see cref="ExcelThreadedComment"/>s on a <see cref="ExcelWorksheet"/>
/// </summary>
public class ExcelWorksheetThreadedComments
{
    internal ExcelWorksheetThreadedComments(ExcelThreadedCommentPersonCollection persons, ExcelWorksheet worksheet)
    {
        this.Persons = persons;
        this._worksheet = worksheet;
        this._package = worksheet._package;
        this._worksheet._threadedCommentsStore = new Core.CellStore.CellStore<int>();
        this.LoadThreads();
    }

    private readonly ExcelWorksheet _worksheet;
    private readonly ExcelPackage _package;
    internal readonly List<ExcelThreadedCommentThread> _threads = new List<ExcelThreadedCommentThread>();
    private readonly List<int> _threadsIndex = new List<int>();
    //internal int _nextId = 0;

    /// <summary>
    /// A collection of persons referenced by the threaded comments.
    /// </summary>
    public ExcelThreadedCommentPersonCollection Persons { get; private set; }

    /// <summary>
    /// An enumerable of the existing <see cref="ExcelThreadedCommentThread"/>s on the <see cref="ExcelWorksheet">worksheet</see>
    /// </summary>
    public IEnumerable<ExcelThreadedCommentThread> Threads => this._threads.Where(x => x != null);

    /// <summary>
    /// Number of <see cref="ExcelThreadedCommentThread"/>s on the <see cref="ExcelWorksheet">worksheet</see> 
    /// </summary>
    public int Count => this._threadsIndex.Count;

    /// <summary>
    /// The raw xml for the threaded comments
    /// </summary>
    public XmlDocument ThreadedCommentsXml { get; private set; }

    private void LoadThreads()
    {
        ZipPackageRelationshipCollection? commentRels = this._worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaThreadedComment);

        foreach (ZipPackageRelationship? commentPart in commentRels)
        {
            Uri? uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
            ZipPackagePart? part = this._package.ZipPackage.GetPart(uri);
            this.ThreadedCommentsXml = new XmlDocument();
            this.ThreadedCommentsXml.PreserveWhitespace = true;
            XmlHelper.LoadXmlSafe(this.ThreadedCommentsXml, part.GetStream());
            this.AddCommentsFromXml();
        }
    }

    private void AddCommentsFromXml()
    {
        foreach (XmlElement node in this.ThreadedCommentsXml.SelectNodes("tc:ThreadedComments/tc:threadedComment", this._worksheet.NameSpaceManager))
        {
            ExcelThreadedComment? comment = new ExcelThreadedComment(node, this._worksheet.NameSpaceManager, this._worksheet.Workbook);
            ExcelCellAddress? cellAddress = comment.CellAddress;
            int i = -1;
            ExcelThreadedCommentThread thread;

            if (this._worksheet._threadedCommentsStore.Exists(cellAddress.Row, cellAddress.Column, ref i))
            {
                thread = this._threads[this._threadsIndex[i]];
            }
            else
            {
                thread = new ExcelThreadedCommentThread(cellAddress, this.ThreadedCommentsXml, this._worksheet);

                lock (this._worksheet._threadedCommentsStore)
                {
                    i = this._threads.Count;
                    this._worksheet._threadedCommentsStore.SetValue(cellAddress.Row, cellAddress.Column, i);
                    this._threadsIndex.Add(i);
                    this._threads.Add(thread);
                }
            }

            comment.Thread = thread;
            thread.AddComment(comment);
        }
    }

    private static void ValidateCellAddress(string cellAddress)
    {
        Require.Argument(cellAddress).IsNotNullOrEmpty("cellAddress");

        if (!ExcelCellBase.IsValidCellAddress(cellAddress))
        {
            throw new ArgumentException(cellAddress + " is not a valid cell address. Use A1 format.");
        }
    }

    /// <summary>
    /// Adds a new <see cref="ExcelThreadedCommentThread"/> to the cell.
    /// </summary>
    /// <param name="cellAddress"></param>
    /// <exception cref="ArgumentException">Thrown if there was an existing <see cref="ExcelThreadedCommentThread"/> in the cell.</exception>
    /// <returns>The new, empty <see cref="ExcelThreadedCommentThread"/></returns>
    public ExcelThreadedCommentThread Add(string cellAddress)
    {
        ValidateCellAddress(cellAddress);

        return this.Add(new ExcelCellAddress(cellAddress));
    }

    /// <summary>
    /// Adds a new <see cref="ExcelThreadedCommentThread"/> to the cell.
    /// </summary>
    /// <param name="cellAddress">The cell address</param>
    /// <returns>The new, empty <see cref="ExcelThreadedCommentThread"/></returns>
    /// <exception cref="ArgumentException">Thrown if there was an existing <see cref="ExcelThreadedCommentThread"/> in the cell.</exception>
    /// <exception cref="InvalidOperationException">If a note/comment exist in the cell</exception>
    public ExcelThreadedCommentThread Add(ExcelCellAddress cellAddress)
    {
        Require.Argument(cellAddress).IsNotNull("cellAddress");

        if (this._worksheet._threadedCommentsStore.Exists(cellAddress.Row, cellAddress.Column))
        {
            throw new ArgumentException("There is an existing threaded comment thread in cell " + cellAddress.Address);
        }

        if (this._worksheet.Comments[cellAddress] != null)
        {
            throw new InvalidOperationException("There is an existing legacy comment/Note in this cell ("
                                                + cellAddress
                                                + "). See the Worksheet.Comments property. Legacy comments and threaded comments cannot reside in the same cell.");
        }

        if (this.ThreadedCommentsXml == null)
        {
            this.ThreadedCommentsXml = new XmlDocument();
            this.ThreadedCommentsXml.PreserveWhitespace = true;

            this.ThreadedCommentsXml
                .LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
        }

        ExcelThreadedCommentThread? thread = new ExcelThreadedCommentThread(cellAddress, this.ThreadedCommentsXml, this._worksheet);
        this._worksheet._threadedCommentsStore.SetValue(cellAddress.Row, cellAddress.Column, this._threads.Count);
        this._threadsIndex.Add(this._threads.Count);
        this._threads.Add(thread);

        return thread;
    }

    /// <summary>
    /// Returns a <see cref="ExcelThreadedCommentThread"/> for the requested <paramref name="cellAddress"/>.
    /// </summary>
    /// <param name="cellAddress">The requested cell address in A1 format</param>
    /// <returns>An existing <see cref="ExcelThreadedCommentThread"/> or null if no thread exists</returns>
    public ExcelThreadedCommentThread this[string cellAddress]
    {
        get
        {
            ValidateCellAddress(cellAddress);

            return this[new ExcelCellAddress(cellAddress)];
        }
    }

    /// <summary>
    /// Returns a <see cref="ExcelThreadedCommentThread"/> for the requested <paramref name="cellAddress"/>.
    /// </summary>
    /// <param name="cellAddress">The requested cell address in A1 format</param>
    /// <returns>An existing <see cref="ExcelThreadedCommentThread"/> or null if no thread exists</returns>
    public ExcelThreadedCommentThread this[ExcelCellAddress cellAddress]
    {
        get
        {
            int i = 0;

            if (this._worksheet._threadedCommentsStore.Exists(cellAddress.Row, cellAddress.Column, ref i))
            {
                return this._threads[i];
            }

            return null;
        }
    }

    /// <summary>
    /// Returns a <see cref="ExcelThreadedCommentThread"/> for the requested <paramref name="index"/>.
    /// </summary>
    /// <param name="index">The index in the collection</param>
    /// <returns>An existing <see cref="ExcelThreadedCommentThread"/> or null if no thread exists</returns>
    public ExcelThreadedCommentThread this[int index]
    {
        get
        {
            if (index < 0 || index >= this._threadsIndex.Count)
            {
                throw new ArgumentOutOfRangeException("Threaded comment index out of range");
            }

            return this._threads[this._threadsIndex[index]];
        }
    }

    /// <summary>
    /// Removes the <see cref="ExcelThreadedCommentThread"/> index position in the collection
    /// </summary>
    /// <param name="index">The index for the threaded comment to be removed</param>
    public void RemoveAt(int index) => this.Remove(this[index]);

    /// <summary>
    /// Removes the <see cref="ExcelThreadedCommentThread"/> supplied
    /// </summary>
    /// <param name="threadedComment">An existing <see cref="ExcelThreadedCommentThread"/> in the worksheet</param>
    public void Remove(ExcelThreadedCommentThread threadedComment)
    {
        int i = -1;
        ExcelThreadedCommentThread c = null;

        if (this._worksheet._threadedCommentsStore.Exists(threadedComment.CellAddress.Row, threadedComment.CellAddress.Column, ref i))
        {
            c = this._threads[i];
        }

        if (threadedComment == c)
        {
            ExcelCellAddress? address = threadedComment.CellAddress;
            ExcelComment? comment = this._worksheet.Comments[address];

            if (comment != null) //Check if the underlaying comment exists.
            {
                this._worksheet.Comments.Remove(comment); //If so, Remove it.
            }

            IEnumerable<XmlNode>? nodes = threadedComment.Comments.Select(x => x.TopNode);

            foreach (XmlNode? node in nodes)
            {
                _ = node.ParentNode.RemoveChild(node); //Remove xml node
            }

            this._worksheet._threadedCommentsStore.Delete(threadedComment.CellAddress.Row, threadedComment.CellAddress.Column, 1, 1, false);
            this._threads[i] = null;
            _ = this._threadsIndex.Remove(i);
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
        List<ExcelThreadedCommentThread> deletedComments = new List<ExcelThreadedCommentThread>();

        foreach (ExcelThreadedCommentThread? threadedComment in this._threads.Where(x => x != null))
        {
            ExcelAddressBase? address = new ExcelAddressBase(threadedComment.CellAddress.Address);

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
                deletedComments.Add(threadedComment);
            }
            else
            {
                threadedComment.CellAddress = new ExcelCellAddress(address.Address);
            }
        }

        foreach (ExcelThreadedCommentThread? comment in deletedComments)
        {
            foreach (ExcelThreadedComment? c in comment.Comments)
            {
                _ = c.TopNode.ParentNode.RemoveChild(c.TopNode);
            }

            int ix = this._threads.IndexOf(comment);
            _ = this._threadsIndex.Remove(ix);
            this._threads[ix] = null;
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
        foreach (ExcelThreadedCommentThread? comment in this._threads.Where(x => x != null))
        {
            ExcelAddressBase? address = new ExcelAddressBase(comment.CellAddress.Address);

            if (rows > 0 && address._fromRow >= fromRow && address._fromCol >= fromCol && address._toCol <= toCol)
            {
                comment.CellAddress = new ExcelCellAddress(address.AddRow(fromRow, rows).Address);
            }

            if (columns > 0 && address._fromCol >= fromCol && address._fromRow >= fromRow && address._toRow <= toRow)
            {
                comment.CellAddress = new ExcelCellAddress(address.AddColumn(fromCol, columns).Address);
            }
        }
    }

    /// <summary>
    ///     Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString() => "Count = " + this._threadsIndex.Count;
}