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

using OfficeOpenXml.Packaging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// Represents a collection of <see cref="ExcelThreadedCommentPerson"/>s in a workbook.
/// </summary>
public class ExcelThreadedCommentPersonCollection : IEnumerable<ExcelThreadedCommentPerson>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="workbook">The <paramref name="workbook"/> where the <see cref="ExcelThreadedCommentPerson"/> occurs</param>
    public ExcelThreadedCommentPersonCollection(ExcelWorkbook workbook)
    {
        this._workbook = workbook;

        if (workbook._package.ZipPackage.PartExists(workbook.PersonsUri))
        {
            this.PersonsXml = workbook._package.GetXmlFromUri(workbook.PersonsUri);

            // lägg upp personerna i listan, loopa på noderna
            XmlElement? listNode = this.PersonsXml.DocumentElement;

            foreach (object? personNode in listNode.ChildNodes)
            {
                ExcelThreadedCommentPerson? person = new ExcelThreadedCommentPerson(workbook.NameSpaceManager, (XmlNode)personNode);
                this._personList.Add(person);
            }
        }
        else
        {
            this.PersonsXml = new XmlDocument();
            this.PersonsXml.LoadXml("<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
        }
    }

    private readonly ExcelWorkbook _workbook;
    private readonly List<ExcelThreadedCommentPerson> _personList = new List<ExcelThreadedCommentPerson>();

    /// <summary>
    /// The xml document defining the threaded comments persons collection.
    /// </summary>
    public XmlDocument PersonsXml { get; private set; }

    /// <summary>
    /// Number of <see cref="ExcelThreadedCommentPerson"/>s in the collection
    /// </summary>
    public int Count => this._personList.Count;

    /// <summary>
    /// Returns the <see cref="ExcelThreadedCommentPerson"/> by its index
    /// </summary>
    /// <param name="index">The requested index</param>
    /// <returns>The <see cref="ExcelThreadedCommentPerson"/> at the requested index</returns>
    public ExcelThreadedCommentPerson this[int index] => this._personList[index];

    /// <summary>
    /// Returns a <see cref="ExcelThreadedCommentPerson"/> by its id
    /// </summary>
    /// <param name="id">The Id of the Person</param>
    /// <returns>A <see cref="ExcelThreadedCommentPerson"/> with the requested <paramref name="id"/> or null</returns>
    public ExcelThreadedCommentPerson this[string id] => this._personList.FirstOrDefault(x => x.Id == id);

    /// <summary>
    /// Finds a <see cref="ExcelThreadedCommentPerson"/> that <paramref name="match"/> a certain criteria
    /// </summary>
    /// <param name="match">The criterias</param>
    /// <returns>A matching <see cref="ExcelThreadedCommentPerson"/></returns>
    public ExcelThreadedCommentPerson Find(Predicate<ExcelThreadedCommentPerson> match) => this._personList.Find(match);

    /// <summary>
    /// Finds a number of <see cref="ExcelThreadedCommentPerson"/>'s that matches a certain criteria.
    /// </summary>
    /// <param name="match">The criterias</param>
    /// <returns>An enumerable of matching <see cref="ExcelThreadedCommentPerson"/>'s</returns>
    public IEnumerable<ExcelThreadedCommentPerson> FindAll(Predicate<ExcelThreadedCommentPerson> match) => this._personList.FindAll(match);

    /// <summary>
    /// Creates and adds a new <see cref="ExcelThreadedCommentPerson"/> to the workbooks list of persons. A unique Id for the person will be generated and set.
    /// The userId will be the same as the <paramref name="displayName"/> and identityProvider will be set to <see cref="IdentityProvider.NoProvider"/>
    /// </summary>
    /// <param name="displayName">The display name of the added <see cref="ExcelThreadedCommentPerson"/></param>
    public ExcelThreadedCommentPerson Add(string displayName) => this.Add(displayName, displayName, IdentityProvider.NoProvider);

    /// <summary>
    /// Creates and adds a new <see cref="ExcelThreadedCommentPerson"/> to the workbooks list of persons. A unique Id for the person will be generated and set.
    /// </summary>
    /// <param name="displayName">The display name of the added <see cref="ExcelThreadedCommentPerson"/></param>
    /// <param name="userId">A string representing the userId of the <paramref name="identityProvider"/></param>
    /// <param name="identityProvider">The <see cref="IdentityProvider"/> from which the <see cref="ExcelThreadedCommentPerson"/> originates</param>
    /// <returns>The added <see cref="ExcelThreadedCommentPerson"/></returns>
    public ExcelThreadedCommentPerson Add(string displayName, string userId, IdentityProvider identityProvider) => this.Add(displayName, userId, identityProvider, ExcelThreadedCommentPerson.NewId());

    /// <summary>
    /// Creates and adds a new <see cref="ExcelThreadedCommentPerson"/> to the workbooks list of persons
    /// </summary>
    /// <param name="displayName">The display name of the added <see cref="ExcelThreadedCommentPerson"/></param>
    /// <param name="userId">A string representing the userId of the <paramref name="identityProvider"/></param>
    /// <param name="identityProvider">The <see cref="IdentityProvider"/> from which the <see cref="ExcelThreadedCommentPerson"/> originates</param>
    /// <param name="id">Id of the <see cref="ExcelThreadedCommentPerson"/></param>
    /// <returns>The added <see cref="ExcelThreadedCommentPerson"/></returns>
    public ExcelThreadedCommentPerson Add(string displayName, string userId, IdentityProvider identityProvider, string id)
    {
        XmlElement? personsNode = this.PersonsXml.CreateElement("person", ExcelPackage.schemaThreadedComments);
        _ = this.PersonsXml.DocumentElement.AppendChild(personsNode);
        ExcelThreadedCommentPerson? p = new ExcelThreadedCommentPerson(this._workbook.NameSpaceManager, personsNode);
        p.DisplayName = displayName;
        p.Id = id;
        p.UserId = userId;
        p.ProviderId = identityProvider;
        this._personList.Add(p);

        return p;
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<ExcelThreadedCommentPerson> GetEnumerator() => this._personList.GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => this._personList.GetEnumerator();

    internal void Add(ExcelThreadedCommentPerson person) => this._personList.Add(person);

    /// <summary>
    /// Removes a <see cref="ExcelThreadedCommentPerson"/> from the collection
    /// </summary>
    /// <param name="person"></param>
    public void Remove(ExcelThreadedCommentPerson person)
    {
        XmlNode? node = this.PersonsXml.DocumentElement.SelectSingleNode("/person[id='" + person.Id + "']");

        if (node != null)
        {
            _ = this.PersonsXml.DocumentElement.RemoveChild(node);
        }

        _ = this._personList.Remove(person);
    }

    /// <summary>
    /// Removes all persons from the collection
    /// </summary>
    public void Clear()
    {
        this.PersonsXml.DocumentElement.RemoveAll();
        this._personList.Clear();
    }

    /// <summary>
    ///     Returns a string that represents the current object.
    /// </summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString() => "Count = " + this._personList.Count;

    internal void Save(ExcelPackage package, ZipPackagePart WorkbookPart, Uri personsUri)
    {
        if (this.Count == 0)
        {
            if (package.ZipPackage.PartExists(personsUri))
            {
                package.ZipPackage.DeletePart(personsUri);
            }
        }
        else
        {
            if (!package.ZipPackage.PartExists(personsUri))
            {
                _ = package.ZipPackage.CreatePart(personsUri, "application/vnd.ms-excel.person+xml");
                _ = WorkbookPart.CreateRelationship(personsUri, TargetMode.Internal, ExcelPackage.schemaPersonsRelationShips);
            }

            package.SavePart(personsUri, this.PersonsXml);
        }
    }
}