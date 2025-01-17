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
using OfficeOpenXml.Utils;
using OfficeOpenXml.Encryption;

namespace OfficeOpenXml;

/// <summary>
/// Sets protection on the workbook level
///<seealso cref="ExcelEncryption"/> 
///<seealso cref="ExcelSheetProtection"/> 
/// </summary>
public class ExcelProtection : XmlHelper
{
    internal ExcelProtection(XmlNamespaceManager ns, XmlNode topNode, ExcelWorkbook wb)
        : base(ns, topNode) =>
        this.SchemaNodeOrder = wb.SchemaNodeOrder;

    const string workbookPasswordPath = "d:workbookProtection/@workbookPassword";

    /// <summary>
    /// Sets a password for the workbook. This does not encrypt the workbook. 
    /// </summary>
    /// <param name="Password">The password. </param>
    public void SetPassword(string Password)
    {
        if (string.IsNullOrEmpty(Password))
        {
            this.DeleteNode(workbookPasswordPath);
        }
        else
        {
            this.SetXmlNodeString(workbookPasswordPath, ((int)EncryptedPackageHandler.CalculatePasswordHash(Password)).ToString("x"));
        }
    }

    const string lockStructurePath = "d:workbookProtection/@lockStructure";

    /// <summary>
    /// Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.
    /// </summary>
    public bool LockStructure
    {
        get => this.GetXmlNodeBool(lockStructurePath, false);
        set => this.SetXmlNodeBool(lockStructurePath, value, false);
    }

    const string lockWindowsPath = "d:workbookProtection/@lockWindows";

    /// <summary>
    /// Locks the position of the workbook window.
    /// </summary>
    public bool LockWindows
    {
        get => this.GetXmlNodeBool(lockWindowsPath, false);
        set => this.SetXmlNodeBool(lockWindowsPath, value, false);
    }

    const string lockRevisionPath = "d:workbookProtection/@lockRevision";

    /// <summary>
    /// Lock the workbook for revision
    /// </summary>
    public bool LockRevision
    {
        get => this.GetXmlNodeBool(lockRevisionPath, false);
        set => this.SetXmlNodeBool(lockRevisionPath, value, false);
    }

    ExcelWriteProtection _writeProtection;

    /// <summary>
    /// File sharing settings for the workbook.
    /// </summary>
    public ExcelWriteProtection WriteProtection => this._writeProtection ??= new ExcelWriteProtection(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
}