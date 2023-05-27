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
using System.Security.Cryptography;

namespace OfficeOpenXml.VBA;

/// <summary>
/// Vba security properties
/// </summary>
public class ExcelVbaProtection
{
    ExcelVbaProject _project;

    internal ExcelVbaProtection(ExcelVbaProject project)
    {
        this._project = project;
        this.VisibilityState = true;
    }

    /// <summary>
    /// If access to the VBA project was restricted by the user
    /// </summary>
    public bool UserProtected { get; internal set; }

    /// <summary>
    /// If access to the VBA project was restricted by the VBA host application
    /// </summary>
    public bool HostProtected { get; internal set; }

    /// <summary>
    /// If access to the VBA project was restricted by the VBA project editor
    /// </summary>
    public bool VbeProtected { get; internal set; }

    /// <summary>
    /// if the VBA project is visible.
    /// </summary>
    public bool VisibilityState { get; internal set; }

    internal byte[] PasswordHash { get; set; }

    internal byte[] PasswordKey { get; set; }

    /// <summary>
    /// Password protect the VBA project.
    /// An empty string or null will remove the password protection
    /// </summary>
    /// <param name="Password">The password</param>
    public void SetPassword(string Password)
    {
        if (string.IsNullOrEmpty(Password))
        {
            this.PasswordHash = null;
            this.PasswordKey = null;
            this.VbeProtected = false;
            this.HostProtected = false;
            this.UserProtected = false;
            this.VisibilityState = true;
            this._project.ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
        }
        else
        {
            //Join Password and Key
            //Set the key
            this.PasswordKey = new byte[4];
            RandomNumberGenerator r = RandomNumberGenerator.Create();
            r.GetBytes(this.PasswordKey);

            byte[] data = new byte[Password.Length + 4];
            Array.Copy(Encoding.GetEncoding(this._project.CodePage).GetBytes(Password), data, Password.Length);
            this.VbeProtected = true;
            this.VisibilityState = false;
            Array.Copy(this.PasswordKey, 0, data, data.Length - 4, 4);

            //Calculate Hash
            SHA1? provider = SHA1.Create();
            this.PasswordHash = provider.ComputeHash(data);
            this._project.ProjectID = "{00000000-0000-0000-0000-000000000000}";
        }
    }

    //public void ValidatePassword(string Password)                     
    //{

    //}        
}