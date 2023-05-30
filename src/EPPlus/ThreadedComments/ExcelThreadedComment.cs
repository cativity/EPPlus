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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments;

/// <summary>
/// Represents a comment in a thread of ThreadedComments
/// </summary>
public class ExcelThreadedComment : XmlHelper
{
    internal ExcelThreadedComment(XmlNode topNode, XmlNamespaceManager namespaceManager, ExcelWorkbook workbook)
        : this(topNode, namespaceManager, workbook, null)
    {
    }

    internal ExcelThreadedComment(XmlNode topNode, XmlNamespaceManager namespaceManager, ExcelWorkbook workbook, ExcelThreadedCommentThread thread)
        : base(namespaceManager, topNode)
    {
        this.SchemaNodeOrder = new string[] { "text", "mentions" };
        this._workbook = workbook;
        this._thread = thread;
    }

    private readonly ExcelWorkbook _workbook;
    private ExcelThreadedCommentThread _thread;

    internal ExcelThreadedCommentThread Thread
    {
        set => this._thread = value ?? throw new ArgumentNullException("Thread");
    }

    internal static string NewId()
    {
        Guid guid = Guid.NewGuid();

        return "{" + guid.ToString().ToUpper() + "}";
    }

    /// <summary>
    /// Indicates whether the Text contains mentions. If so the
    /// Mentions property will contain data about those mentions.
    /// </summary>
    public bool ContainsMentions => this.Mentions != null && this.Mentions.Any();

    /// <summary>
    /// Address of the cell in the A1 format
    /// </summary>
    internal string Ref
    {
        get => this.GetXmlNodeString("@ref");
        set => this.SetXmlNodeString("@ref", value);
    }

    private ExcelCellAddress _cellAddress;

    /// <summary>
    /// The location of the threaded comment
    /// </summary>
    public ExcelCellAddress CellAddress
    {
        get => this._cellAddress ??= new ExcelCellAddress(this.Ref);
        internal set
        {
            this._cellAddress = value;
            this.Ref = this.CellAddress.Address;
        }
    }

    /// <summary>
    /// Timestamp for when the comment was created
    /// </summary>
    public DateTime DateCreated
    {
        get
        {
            string? dt = this.GetXmlNodeString("@dT");

            if (DateTime.TryParse(dt, out DateTime result))
            {
                return result;
            }

            throw new InvalidCastException("Could not cast datetime for threaded comment");
        }
        set => this.SetXmlNodeString("@dT", value.ToString("yyyy-MM-ddTHH:mm:ss.ff"));
    }

    /// <summary>
    /// Unique id
    /// </summary>
    public string Id
    {
        get => this.GetXmlNodeString("@id");
        internal set => this.SetXmlNodeString("@id", value);
    }

    /// <summary>
    /// Id of the <see cref="ExcelThreadedCommentPerson"/> who wrote the comment
    /// </summary>
    public string PersonId
    {
        get => this.GetXmlNodeString("@personId");
        set
        {
            this.SetXmlNodeString("@personId", value);
            this._thread.OnCommentThreadChanged();
        }
    }

    /// <summary>
    /// Author of the comment
    /// </summary>
    public ExcelThreadedCommentPerson Author => this._workbook.ThreadedCommentPersons[this.PersonId];

    /// <summary>
    /// Id of the first comment in the thread
    /// </summary>
    public string ParentId
    {
        get => this.GetXmlNodeString("@parentId");
        set
        {
            this.SetXmlNodeString("@parentId", value);
            this._thread.OnCommentThreadChanged();
        }
    }

    internal bool? Done
    {
        get
        {
            string? val = this.GetXmlNodeString("@done");

            if (string.IsNullOrEmpty(val))
            {
                return null;
            }

            if (val == "1")
            {
                return true;
            }

            return false;
        }
        set
        {
            if (value.HasValue && value.Value)
            {
                this.SetXmlNodeInt("@done", 1);
            }
            else if (value.HasValue && !value.Value)
            {
                this.SetXmlNodeInt("@done", 0);
            }
            else
            {
                this.SetXmlNodeInt("@done", null);
            }
        }
    }

    /// <summary>
    /// Text of the comment. To edit the text on an existing comment, use the EditText function.
    /// </summary>
    public string Text
    {
        get => this.GetXmlNodeString("tc:text");
        internal set
        {
            this.SetXmlNodeString("tc:text", value);
            this._thread.OnCommentThreadChanged();
        }
    }

    /// <summary>
    /// Edit the Text of an existing comment
    /// </summary>
    /// <param name="newText"></param>
    public void EditText(string newText)
    {
        this.Mentions.Clear();
        this.Text = newText;
        this._thread.OnCommentThreadChanged();
    }

    /// <summary>
    /// Edit the Text of an existing comment with mentions
    /// </summary>
    /// <param name="newTextWithFormats">A string with format placeholders - same as in string.Format. Index in these should correspond to an index in the <paramref name="personsToMention"/> array.</param>
    /// <param name="personsToMention">A params array of <see cref="ExcelThreadedCommentPerson"/>. Their DisplayName property will be used to replace the format placeholders.</param>
    public void EditText(string newTextWithFormats, params ExcelThreadedCommentPerson[] personsToMention)
    {
        this.Mentions.Clear();
        MentionsHelper.InsertMentions(this, newTextWithFormats, personsToMention);
        this._thread.OnCommentThreadChanged();
    }

    private ExcelThreadedCommentMentionCollection _mentions;

    /// <summary>
    /// Mentions in this comment. Will return null if no mentions exists.
    /// </summary>
    public ExcelThreadedCommentMentionCollection Mentions
    {
        get
        {
            if (this._mentions == null)
            {
                XmlNode? mentionsNode = this.TopNode.SelectSingleNode("tc:mentions", this.NameSpaceManager);

                if (mentionsNode == null)
                {
                    mentionsNode = this.TopNode.OwnerDocument.CreateElement("mentions", ExcelPackage.schemaThreadedComments);
                    _ = this.TopNode.AppendChild(mentionsNode);
                }

                this._mentions = new ExcelThreadedCommentMentionCollection(this.NameSpaceManager, mentionsNode);
            }

            return this._mentions;
        }
    }
}