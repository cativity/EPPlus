﻿/*************************************************************************************************
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

using OfficeOpenXml.Encryption;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Security.Cryptography;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// File sharing settings for the workbook.
/// </summary>
public class ExcelWriteProtection : XmlHelper
{
    internal ExcelWriteProtection(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder)
        : base(nameSpaceManager, topNode) =>
        this.SchemaNodeOrder = schemaNodeOrder;

    /// <summary>
    /// Writes protectes the workbook with a password. 
    /// EPPlus uses SHA-512 as hash algorithm with a spin count of 100000.
    /// </summary>
    /// <param name="userName">The name of the person enforcing the write protection</param>
    /// <param name="password">The password. Setting the password to null or empty will remove the read-only mode.</param>
    public void SetReadOnly(string userName, string password)
    {
        this.UserName = userName;

        if (string.IsNullOrEmpty(password?.Trim()))
        {
            this.RemovePasswordAttributes();

            return;
        }

        this.HashAlgorithm = eHashAlgorithm.SHA512;

        byte[]? s = new byte[16];
        RandomNumberGenerator? rnd = RandomNumberGenerator.Create();
        rnd.GetBytes(s);
        this.SaltValue = s;
        this.SpinCount = 100000;

        this.HashValue = EncryptedPackageHandler.GetPasswordHashSpinAppending(SHA512.Create(), this.SaltValue, password, this.SpinCount, 64);
    }

    private void RemovePasswordAttributes()
    {
        XmlElement? node = (XmlElement)this.GetNode("d:fileSharing");
        node.RemoveAttribute("spinCount");
        node.RemoveAttribute("saltValue");
        node.RemoveAttribute("hashValue");
    }

    /// <summary>
    /// Remove any write protection set on the workbook
    /// </summary>
    public void RemoveReadOnly() => this.DeleteNode("d:fileSharing");

    internal eHashAlgorithm HashAlgorithm
    {
        get => GetHashAlogorithm(this.GetXmlNodeString("d:fileSharing/@algorithmName"));
        private set => this.SetXmlNodeString("d:fileSharing/@algorithmName", SetHashAlogorithm(value));
    }

    private static string SetHashAlogorithm(eHashAlgorithm value)
    {
        switch (value)
        {
            case eHashAlgorithm.SHA512:
                return "SHA-512";

            default:
                throw new NotSupportedException("EPPlus only support SHA 512 hashing for file sharing");
        }
    }

    private static eHashAlgorithm GetHashAlogorithm(string v)
    {
        switch (v)
        {
            case "RIPEMD-128":
                return eHashAlgorithm.RIPEMD128;

            case "RIPEMD-160":
                return eHashAlgorithm.RIPEMD160;

            case "SHA-1":
                return eHashAlgorithm.SHA1;

            case "SHA-256":
                return eHashAlgorithm.SHA256;

            case "SHA-384":
                return eHashAlgorithm.SHA384;

            case "SHA-512":
                return eHashAlgorithm.SHA512;

            default:
                return v.ToEnum(eHashAlgorithm.SHA512);
        }
    }

    internal int SpinCount
    {
        get => this.GetXmlNodeInt("d:fileSharing/@spinCount");
        set => this.SetXmlNodeInt("d:fileSharing/@spinCount", value);
    }

    internal byte[] SaltValue
    {
        get
        {
            string? s = this.GetXmlNodeString("d:fileSharing/@saltValue");

            if (!string.IsNullOrEmpty(s))
            {
                return Convert.FromBase64String(s);
            }

            return null;
        }
        set => this.SetXmlNodeString("d:fileSharing/@saltValue", Convert.ToBase64String(value));
    }

    internal byte[] HashValue
    {
        get
        {
            string? s = this.GetXmlNodeString("d:fileSharing/@hashValue");

            if (!string.IsNullOrEmpty(s))
            {
                return Convert.FromBase64String(s);
            }

            return null;
        }
        set => this.SetXmlNodeString("d:fileSharing/@hashValue", Convert.ToBase64String(value));
    }

    /// <summary>
    /// If the workbook is set to readonly and has a password set.
    /// </summary>
    public bool IsReadOnly => this.ExistsNode("d:fileSharing/@hashValue");

    /// <summary>
    /// The name of the person enforcing the write protection.
    /// </summary>
    public string UserName
    {
        get => this.GetXmlNodeString("d:fileSharing/@userName");
        set => this.SetXmlNodeString("d:fileSharing/@userName", value);
    }

    /// <summary>
    /// If the author recommends that you open the workbook in read-only mode.
    /// </summary>
    public bool ReadOnlyRecommended
    {
        get => this.GetXmlNodeBool("d:fileSharing/@readOnlyRecommended");
        set => this.SetXmlNodeBool("d:fileSharing/@readOnlyRecommended", value);
    }
}