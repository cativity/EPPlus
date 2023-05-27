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

using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Encryption;

internal abstract class EncryptionInfo
{
    internal short MajorVersion;
    internal short MinorVersion;

    internal abstract void Read(byte[] data);

    internal static EncryptionInfo ReadBinary(byte[] data)
    {
        short majorVersion = BitConverter.ToInt16(data, 0);
        short minorVersion = BitConverter.ToInt16(data, 2);
        EncryptionInfo ret;

        if ((minorVersion == 2 || minorVersion == 3) && majorVersion <= 4) // minorVersion==1 is RC4, not supported.
        {
            ret = new EncryptionInfoBinary();
        }
        else if (majorVersion == 4 && minorVersion == 4)
        {
            ret = new EncryptionInfoAgile();
        }
        else
        {
            throw new NotSupportedException("Unsupported encryption format");
        }

        ret.MajorVersion = majorVersion;
        ret.MinorVersion = minorVersion;
        ret.Read(data);

        return ret;
    }
}

internal enum eCipherAlgorithm
{
    /// <summary>
    /// AES. MUST conform to the AES algorithm.
    /// </summary>
    AES,

    /// <summary>
    /// RC2. MUST conform to [RFC2268].
    /// </summary>
    RC2,

    /// <summary>
    /// RC4. 
    /// </summary>
    RC4,

    /// <summary>
    /// MUST conform to the DES algorithm.
    /// </summary>
    DES,

    /// <summary>
    /// MUST conform to the [DRAFT-DESX] algorithm.
    /// </summary>
    DESX,

    /// <summary>
    /// 3DES. MUST conform to the [RFC1851] algorithm. 
    /// </summary>
    TRIPLE_DES,

    /// 3DES_112 MUST conform to the [RFC1851] algorithm. 
    TRIPLE_DES_112
}

internal enum eChainingMode
{
    /// <summary>
    /// Cipher block chaining (CBC).
    /// </summary>
    ChainingModeCBC,

    /// <summary>
    /// Cipher feedback chaining (CFB), with 8-bit window.
    /// </summary>
    ChainingModeCFB
}

/// <summary>
/// Hash algorithm
/// </summary>
internal enum eHashAlgorithm
{
    /// <summary>
    /// Sha 1-MUST conform to [RFC4634]
    /// </summary>
    SHA1,

    /// <summary>
    /// Sha 256-MUST conform to [RFC4634]
    /// </summary>
    SHA256,

    /// <summary>
    /// Sha 384-MUST conform to [RFC4634]
    /// </summary>
    SHA384,

    /// <summary>
    /// Sha 512-MUST conform to [RFC4634]
    /// </summary>
    SHA512,

    /// <summary>
    /// MD5
    /// </summary>
    MD5,

    /// <summary>
    /// MD4
    /// </summary>
    MD4,

    /// <summary>
    /// MD2
    /// </summary>
    MD2,

    /// <summary>
    /// RIPEMD-128 MUST conform to [ISO/IEC 10118]
    /// </summary>
    RIPEMD128,

    /// <summary>
    /// RIPEMD-160 MUST conform to [ISO/IEC 10118]
    /// </summary>
    RIPEMD160,

    /// <summary>
    /// WHIRLPOOL MUST conform to [ISO/IEC 10118]
    /// </summary>
    WHIRLPOOL
}

/// <summary>
/// Handels the agile encryption
/// </summary>
internal class EncryptionInfoAgile : EncryptionInfo
{
    XmlNamespaceManager _nsm;

    public EncryptionInfoAgile()
    {
        NameTable? nt = new NameTable();
        this._nsm = new XmlNamespaceManager(nt);
        this._nsm.AddNamespace("d", "http://schemas.microsoft.com/office/2006/encryption");
        this._nsm.AddNamespace("c", "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");
        this._nsm.AddNamespace("p", "http://schemas.microsoft.com/office/2006/keyEncryptor/password");
    }

    internal class EncryptionKeyData : XmlHelper
    {
        public EncryptionKeyData(XmlNamespaceManager nsm, XmlNode topNode)
            : base(nsm, topNode)
        {
        }

        internal byte[] SaltValue
        {
            get
            {
                string? s = this.GetXmlNodeString("@saltValue");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@saltValue", Convert.ToBase64String(value)); }
        }

        internal eHashAlgorithm HashAlgorithm
        {
            get { return GetHashAlgorithm(this.GetXmlNodeString("@hashAlgorithm")); }
            set { this.SetXmlNodeString("@hashAlgorithm", GetHashAlgorithmString(value)); }
        }

        private static eHashAlgorithm GetHashAlgorithm(string v)
        {
            switch (v)
            {
                case "RIPEMD-128":
                    return eHashAlgorithm.RIPEMD128;

                case "RIPEMD-160":
                    return eHashAlgorithm.RIPEMD160;

                case "SHA-1":
                    return eHashAlgorithm.SHA1;

                default:
                    try
                    {
                        return (eHashAlgorithm)Enum.Parse(typeof(eHashAlgorithm), v);
                    }
                    catch
                    {
                        throw new InvalidDataException("Invalid Hash algorithm");
                    }
            }
        }

        private static string GetHashAlgorithmString(eHashAlgorithm value)
        {
            switch (value)
            {
                case eHashAlgorithm.RIPEMD128:
                    return "RIPEMD-128";

                case eHashAlgorithm.RIPEMD160:
                    return "RIPEMD-160";

                case eHashAlgorithm.SHA1:
                    return "SHA-1";

                default:
                    return value.ToString();
            }
        }

        internal eChainingMode CipherChaining
        {
            get
            {
                string? v = this.GetXmlNodeString("@cipherChaining");

                try
                {
                    return (eChainingMode)Enum.Parse(typeof(eChainingMode), v);
                }
                catch
                {
                    throw new InvalidDataException("Invalid chaining mode");
                }
            }
            set { this.SetXmlNodeString("@cipherChaining", value.ToString()); }
        }

        internal eCipherAlgorithm CipherAlgorithm
        {
            get { return GetCipherAlgorithm(this.GetXmlNodeString("@cipherAlgorithm")); }
            set { this.SetXmlNodeString("@cipherAlgorithm", GetCipherAlgorithmString(value)); }
        }

        private static eCipherAlgorithm GetCipherAlgorithm(string v)
        {
            switch (v)
            {
                case "3DES":
                    return eCipherAlgorithm.TRIPLE_DES;

                case "3DES_112":
                    return eCipherAlgorithm.TRIPLE_DES_112;

                default:
                    try
                    {
                        return (eCipherAlgorithm)Enum.Parse(typeof(eCipherAlgorithm), v);
                    }
                    catch
                    {
                        throw new InvalidDataException("Invalid Hash algorithm");
                    }
            }
        }

        private static string GetCipherAlgorithmString(eCipherAlgorithm alg)
        {
            switch (alg)
            {
                case eCipherAlgorithm.TRIPLE_DES:
                    return "3DES";

                case eCipherAlgorithm.TRIPLE_DES_112:
                    return "3DES_112";

                default:
                    return alg.ToString();
            }
        }

        internal int HashSize
        {
            get { return this.GetXmlNodeInt("@hashSize"); }
            set { this.SetXmlNodeString("@hashSize", value.ToString()); }
        }

        internal int KeyBits
        {
            get { return this.GetXmlNodeInt("@keyBits"); }
            set { this.SetXmlNodeString("@keyBits", value.ToString()); }
        }

        internal int BlockSize
        {
            get { return this.GetXmlNodeInt("@blockSize"); }
            set { this.SetXmlNodeString("@blockSize", value.ToString()); }
        }

        internal int SaltSize
        {
            get { return this.GetXmlNodeInt("@saltSize"); }
            set { this.SetXmlNodeString("@saltSize", value.ToString()); }
        }
    }

    internal class EncryptionDataIntegrity : XmlHelper
    {
        public EncryptionDataIntegrity(XmlNamespaceManager nsm, XmlNode topNode)
            : base(nsm, topNode)
        {
        }

        internal byte[] EncryptedHmacValue
        {
            get
            {
                string? s = this.GetXmlNodeString("@encryptedHmacValue");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@encryptedHmacValue", Convert.ToBase64String(value)); }
        }

        internal byte[] EncryptedHmacKey
        {
            get
            {
                string? s = this.GetXmlNodeString("@encryptedHmacKey");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@encryptedHmacKey", Convert.ToBase64String(value)); }
        }
    }

    internal class EncryptionKeyEncryptor : EncryptionKeyData
    {
        public EncryptionKeyEncryptor(XmlNamespaceManager nsm, XmlNode topNode)
            : base(nsm, topNode)
        {
        }

        internal byte[] EncryptedKeyValue
        {
            get
            {
                string? s = this.GetXmlNodeString("@encryptedKeyValue");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@encryptedKeyValue", Convert.ToBase64String(value)); }
        }

        internal byte[] EncryptedVerifierHash
        {
            get
            {
                string? s = this.GetXmlNodeString("@encryptedVerifierHashValue");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@encryptedVerifierHashValue", Convert.ToBase64String(value)); }
        }

        internal byte[] EncryptedVerifierHashInput
        {
            get
            {
                string? s = this.GetXmlNodeString("@encryptedVerifierHashInput");

                if (!string.IsNullOrEmpty(s))
                {
                    return Convert.FromBase64String(s);
                }

                return null;
            }
            set { this.SetXmlNodeString("@encryptedVerifierHashInput", Convert.ToBase64String(value)); }
        }

        internal byte[] VerifierHashInput { get; set; }

        internal byte[] VerifierHash { get; set; }

        internal byte[] KeyValue { get; set; }

        internal int SpinCount
        {
            get { return this.GetXmlNodeInt("@spinCount"); }
            set { this.SetXmlNodeString("@spinCount", value.ToString()); }
        }
    }
    /*
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
       <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
           <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="pa+hrJ3s1zrY6hmVuSa5JQ==" />
           <dataIntegrity encryptedHmacKey="nd8i4sEKjsMjVN2gLo91oFN2e7bhMpWKDCAUBEpz4GW6NcE3hBXDobLksZvQGwLrPj0SUVzQA8VuDMyjMAfVCA==" encryptedHmacValue="O6oegHpQVz2uO7Om4oZijSi4kzLiiMZGIjfZlq/EFFO6PZbKitenBqe2or1REaxaI7gO/JmtJzZ1ViucqTaw4g==" />
           <keyEncryptors>
               <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="u2BNFAuHYn3M/WRja3/uPg==" encryptedVerifierHashInput="M0V+fRolJMRgFyI9w+AVxQ==" encryptedVerifierHashValue="V/6l9pFH7AaXFqEbsnFBfHe7gMOqFeRwaNMjc7D3LNdw6KgZzOOQlt5sE8/oG7GPVBDGfoQMTxjQydVPVy4qng==" encryptedKeyValue="B0/rbSQRiIKG5CQDH6AKYSybdXzxgKAfX1f+S5k7mNE=" />
               </keyEncryptor></keyEncryptors></encryption>
    */

    /***
     * <?xml version="1.0" encoding="UTF-8" standalone="true"?>
        <encryption xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns="http://schemas.microsoft.com/office/2006/encryption">
     *      <keyData saltValue="XmTB/XBGJSbwd/GTKzQv5A==" hashAlgorithm="SHA512" cipherChaining="ChainingModeCBC" cipherAlgorithm="AES" hashSize="64" keyBits="256" blockSize="16" saltSize="16"/>
     *      <dataIntegrity encryptedHmacValue="WWw3Bb2dbcNPMnl9f1o7rO0u7sclWGKTXqBA6rRzKsP2KzWS5T0LxY9qFoC6QE67t/t+FNNtMDdMtE3D1xvT8w==" encryptedHmacKey="p/dVdlJY5Kj0k3jI1HRjqtk4s0Y4HmDAsc8nqZgfxNS7DopAsS3LU/2p3CYoIRObHsnHTAtbueH08DFCYGZURg=="/>
     *          <keyEncryptors>
     *              <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
     *                  <p:encryptedKey saltValue="EeBtY0QftyOkLztCl7NF0g==" hashAlgorithm="SHA512" cipherChaining="ChainingModeCBC" cipherAlgorithm="AES" hashSize="64" keyBits="256" blockSize="16" saltSize="16" encryptedKeyValue="Z7AO8vHnnPZEb1VqyZLJ6JFc3Mq3E322XPxWXS21fbU=" encryptedVerifierHashValue="G7BxbKnZanldvtsbu51mP9J3f9Wr5vCfCpvWSh5eIJff7Sr3J2DzH1/9aKj9uIpqFQIsLohpRk+oBYDcX7hRgw==" encryptedVerifierHashInput="851eszl5y5rdU1RnTjEWHw==" spinCount="100000"/>
     *              </keyEncryptor>
     *      </keyEncryptors>
     *      </encryption
     * ***/
    internal EncryptionDataIntegrity DataIntegrity { get; set; }

    internal EncryptionKeyData KeyData { get; set; }

    internal List<EncryptionKeyEncryptor> KeyEncryptors { get; private set; }

    internal XmlDocument Xml { get; set; }

    internal override void Read(byte[] data)
    {
        byte[]? byXml = new byte[data.Length - 8];
        Array.Copy(data, 8, byXml, 0, data.Length - 8);
        string? xml = Encoding.UTF8.GetString(byXml);
        this.ReadFromXml(xml);
    }

    internal void ReadFromXml(string xml)
    {
        this.Xml = new XmlDocument();
        XmlHelper.LoadXmlSafe(this.Xml, xml, Encoding.UTF8);
        XmlNode? node = this.Xml.SelectSingleNode("/d:encryption/d:keyData", this._nsm);
        this.KeyData = new EncryptionKeyData(this._nsm, node);
        node = this.Xml.SelectSingleNode("/d:encryption/d:dataIntegrity", this._nsm);
        this.DataIntegrity = new EncryptionDataIntegrity(this._nsm, node);
        this.KeyEncryptors = new List<EncryptionKeyEncryptor>();

        XmlNodeList? list = this.Xml.SelectNodes("/d:encryption/d:keyEncryptors/d:keyEncryptor/p:encryptedKey", this._nsm);

        if (list != null)
        {
            foreach (XmlNode n in list)
            {
                this.KeyEncryptors.Add(new EncryptionKeyEncryptor(this._nsm, n));
            }
        }
    }
}

/// <summary>
/// Handles the EncryptionInfo stream
/// </summary>
internal class EncryptionInfoBinary : EncryptionInfo
{
    internal Flags Flags;
    internal uint HeaderSize;
    internal EncryptionHeader Header;
    internal EncryptionVerifier Verifier;

    internal override void Read(byte[] data)
    {
        this.Flags = (Flags)BitConverter.ToInt32(data, 4);
        this.HeaderSize = (uint)BitConverter.ToInt32(data, 8);

        /**** EncryptionHeader ****/
        this.Header = new EncryptionHeader();
        this.Header.Flags = (Flags)BitConverter.ToInt32(data, 12);
        this.Header.SizeExtra = BitConverter.ToInt32(data, 16);
        this.Header.AlgID = (AlgorithmID)BitConverter.ToInt32(data, 20);
        this.Header.AlgIDHash = (AlgorithmHashID)BitConverter.ToInt32(data, 24);
        this.Header.KeySize = BitConverter.ToInt32(data, 28);
        this.Header.ProviderType = (ProviderType)BitConverter.ToInt32(data, 32);
        this.Header.Reserved1 = BitConverter.ToInt32(data, 36);
        this.Header.Reserved2 = BitConverter.ToInt32(data, 40);

        byte[] text = new byte[(int)this.HeaderSize - 34];
        Array.Copy(data, 44, text, 0, (int)this.HeaderSize - 34);
        this.Header.CSPName = Encoding.Unicode.GetString(text);

        int pos = (int)this.HeaderSize + 12;

        /**** EncryptionVerifier ****/
        this.Verifier = new EncryptionVerifier();
        this.Verifier.SaltSize = (uint)BitConverter.ToInt32(data, pos);
        this.Verifier.Salt = new byte[this.Verifier.SaltSize];

        Array.Copy(data, pos + 4, this.Verifier.Salt, 0, (int)this.Verifier.SaltSize);

        this.Verifier.EncryptedVerifier = new byte[16];
        Array.Copy(data, pos + 20, this.Verifier.EncryptedVerifier, 0, 16);

        this.Verifier.VerifierHashSize = (uint)BitConverter.ToInt32(data, pos + 36);
        this.Verifier.EncryptedVerifierHash = new byte[this.Verifier.VerifierHashSize];
        Array.Copy(data, pos + 40, this.Verifier.EncryptedVerifierHash, 0, (int)this.Verifier.VerifierHashSize);
    }

    internal byte[] WriteBinary()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter bw = new BinaryWriter(ms);

        bw.Write(this.MajorVersion);
        bw.Write(this.MinorVersion);
        bw.Write((int)this.Flags);
        byte[] header = this.Header.WriteBinary();
        bw.Write((uint)header.Length);
        bw.Write(header);
        bw.Write(this.Verifier.WriteBinary());

        bw.Flush();

        return ms.ToArray();
    }
}