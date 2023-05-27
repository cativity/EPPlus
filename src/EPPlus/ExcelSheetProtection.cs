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
using System.Security.Cryptography;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Encryption;

namespace OfficeOpenXml;

/// <summary>
/// Sheet protection
///<seealso cref="ExcelEncryption"/> 
///<seealso cref="ExcelProtection"/> 
/// </summary>
public sealed class ExcelSheetProtection : XmlHelper
{
    internal ExcelSheetProtection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
        : base(nsm, topNode)
    {
        this.SchemaNodeOrder = ws.SchemaNodeOrder;
    }

    bool _hasSheetProtection = false;
    private const string _isProtectedPath = "d:sheetProtection/@sheet";

    /// <summary>
    /// If the worksheet is protected.
    /// </summary>
    public bool IsProtected
    {
        get { return this.GetXmlNodeBool(_isProtectedPath, false); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_isProtectedPath, value, false);

            if (!value)
            {
                this.DeleteAllNode(_isProtectedPath); //delete the whole sheetprotection node
                this._hasSheetProtection = false;
            }
        }
    }

    private void CreatedDefaultNode()
    {
        if (this._hasSheetProtection = false && !this.ExistsNode("d:sheetProtection"))
        {
            this.AllowEditObject = true;
            this.AllowEditScenarios = true;
            this._hasSheetProtection = true;
        }
    }

    private const string _allowSelectLockedCellsPath = "d:sheetProtection/@selectLockedCells";

    /// <summary>
    /// Allow users to select locked cells
    /// </summary>
    public bool AllowSelectLockedCells
    {
        get { return !this.GetXmlNodeBool(_allowSelectLockedCellsPath, false); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowSelectLockedCellsPath, !value, false);
        }
    }

    private const string _allowSelectUnlockedCellsPath = "d:sheetProtection/@selectUnlockedCells";

    /// <summary>
    /// Allow users to select unlocked cells
    /// </summary>
    public bool AllowSelectUnlockedCells
    {
        get { return !this.GetXmlNodeBool(_allowSelectUnlockedCellsPath, false); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowSelectUnlockedCellsPath, !value, false);
        }
    }

    private const string _allowObjectPath = "d:sheetProtection/@objects";

    /// <summary>
    /// Allow users to edit objects
    /// </summary>
    public bool AllowEditObject
    {
        get { return !this.GetXmlNodeBool(_allowObjectPath, false); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowObjectPath, !value, false);
        }
    }

    private const string _allowScenariosPath = "d:sheetProtection/@scenarios";

    /// <summary>
    /// Allow users to edit senarios
    /// </summary>
    public bool AllowEditScenarios
    {
        get { return !this.GetXmlNodeBool(_allowScenariosPath, false); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowScenariosPath, !value, false);
        }
    }

    private const string _allowFormatCellsPath = "d:sheetProtection/@formatCells";

    /// <summary>
    /// Allow users to format cells
    /// </summary>
    public bool AllowFormatCells
    {
        get { return !this.GetXmlNodeBool(_allowFormatCellsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowFormatCellsPath, !value, true);
        }
    }

    private const string _allowFormatColumnsPath = "d:sheetProtection/@formatColumns";

    /// <summary>
    /// Allow users to Format columns
    /// </summary>
    public bool AllowFormatColumns
    {
        get { return !this.GetXmlNodeBool(_allowFormatColumnsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowFormatColumnsPath, !value, true);
        }
    }

    private const string _allowFormatRowsPath = "d:sheetProtection/@formatRows";

    /// <summary>
    /// Allow users to Format rows
    /// </summary>
    public bool AllowFormatRows
    {
        get { return !this.GetXmlNodeBool(_allowFormatRowsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowFormatRowsPath, !value, true);
        }
    }

    private const string _allowInsertColumnsPath = "d:sheetProtection/@insertColumns";

    /// <summary>
    /// Allow users to insert columns
    /// </summary>
    public bool AllowInsertColumns
    {
        get { return !this.GetXmlNodeBool(_allowInsertColumnsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowInsertColumnsPath, !value, true);
        }
    }

    private const string _allowInsertRowsPath = "d:sheetProtection/@insertRows";

    /// <summary>
    /// Allow users to Format rows
    /// </summary>
    public bool AllowInsertRows
    {
        get { return !this.GetXmlNodeBool(_allowInsertRowsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowInsertRowsPath, !value, true);
        }
    }

    private const string _allowInsertHyperlinksPath = "d:sheetProtection/@insertHyperlinks";

    /// <summary>
    /// Allow users to insert hyperlinks
    /// </summary>
    public bool AllowInsertHyperlinks
    {
        get { return !this.GetXmlNodeBool(_allowInsertHyperlinksPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowInsertHyperlinksPath, !value, true);
        }
    }

    private const string _allowDeleteColumns = "d:sheetProtection/@deleteColumns";

    /// <summary>
    /// Allow users to delete columns
    /// </summary>
    public bool AllowDeleteColumns
    {
        get { return !this.GetXmlNodeBool(_allowDeleteColumns, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowDeleteColumns, !value, true);
        }
    }

    private const string _allowDeleteRowsPath = "d:sheetProtection/@deleteRows";

    /// <summary>
    /// Allow users to delete rows
    /// </summary>  
    public bool AllowDeleteRows
    {
        get { return !this.GetXmlNodeBool(_allowDeleteRowsPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowDeleteRowsPath, !value, true);
        }
    }

    private const string _allowSortPath = "d:sheetProtection/@sort";

    /// <summary>
    /// Allow users to sort a range
    /// </summary>
    public bool AllowSort
    {
        get { return !this.GetXmlNodeBool(_allowSortPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowSortPath, !value, true);
        }
    }

    private const string _allowAutoFilterPath = "d:sheetProtection/@autoFilter";

    /// <summary>
    /// Allow users to use autofilters
    /// </summary>
    public bool AllowAutoFilter
    {
        get { return !this.GetXmlNodeBool(_allowAutoFilterPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowAutoFilterPath, !value, true);
        }
    }

    private const string _allowPivotTablesPath = "d:sheetProtection/@pivotTables";

    /// <summary>
    /// Allow users to use pivottables
    /// </summary>
    public bool AllowPivotTables
    {
        get { return !this.GetXmlNodeBool(_allowPivotTablesPath, true); }
        set
        {
            this.CreatedDefaultNode();
            this.SetXmlNodeBool(_allowPivotTablesPath, !value, true);
        }
    }

    private const string _passwordPath = "d:sheetProtection/@password";

    /// <summary>
    /// Sets a password for the sheet.
    /// </summary>
    /// <param name="Password"></param>
    public void SetPassword(string Password)
    {
        if (this.IsProtected == false)
        {
            this.IsProtected = true;
        }

        Password = Password.Trim();

        if (Password == "")
        {
            XmlNode? node = this.TopNode.SelectSingleNode(_passwordPath, this.NameSpaceManager);

            if (node != null)
            {
                _ = (node as XmlAttribute).OwnerElement.Attributes.Remove(node as XmlAttribute);
            }

            return;
        }

        int hash = EncryptedPackageHandler.CalculatePasswordHash(Password);
        this.SetXmlNodeString(_passwordPath, ((int)hash).ToString("x"));
    }
}