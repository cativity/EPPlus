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
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// A protected range in a worksheet
    ///<seealso cref="ExcelProtection"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    ///<seealso cref="ExcelEncryption"/> 
    /// </summary>
    public class ExcelProtectedRange : XmlHelper
    {
        internal ExcelProtectedRange(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }

        /// <summary>
        /// The name of the protected range
        /// </summary>
        public string Name
        {
            get
            {
                return this.GetXmlNodeString("@name");
            }
            set
            {
                this.SetXmlNodeString("@name",value);
            }
        }
        ExcelAddress _address=null;
        /// <summary>
        /// The address of the protected range
        /// </summary>
        public ExcelAddress Address 
        { 
            get
            {
                if(this._address==null)
                {
                    this._address=new ExcelAddress(this.GetXmlNodeString("@sqref"));
                }
                return this._address;
            }
            set
            {
                this.SetXmlNodeString("@sqref", SqRefUtility.ToSqRefAddress(value.Address));
                this._address=value;
            }
        }
        /// <summary>
        /// Sets the password for the range
        /// </summary>
        /// <param name="password">The password used to generete the hash</param>
        public void SetPassword(string password)
        {
            byte[]? byPwd = Encoding.Unicode.GetBytes(password);
            RandomNumberGenerator? rnd = RandomNumberGenerator.Create();
            byte[]? bySalt=new byte[16];
            rnd.GetBytes(bySalt);
            
            //Default SHA512 and 10000 spins
            this.Algorithm=eProtectedRangeAlgorithm.SHA512;
            this.SpinCount = this.SpinCount < 100000 ? 100000 : this.SpinCount;

            //Combine salt and password and calculate the initial hash
#if Core 
            SHA512? hp = SHA512.Create();
#else
            var hp=new SHA512CryptoServiceProvider();
#endif
            byte[]? buffer =new byte[byPwd.Length + bySalt.Length];
            Array.Copy(bySalt, buffer, bySalt.Length);
            Array.Copy(byPwd, 0, buffer, 16, byPwd.Length);
            byte[]? hash = hp.ComputeHash(buffer);

            //Now iterate the number of spinns.
            for (int i = 0; i < this.SpinCount; i++)
            {
                buffer=new byte[hash.Length+4];
                Array.Copy(hash, buffer, hash.Length);
                Array.Copy(BitConverter.GetBytes(i), 0, buffer, hash.Length, 4);
                hash = hp.ComputeHash(buffer);
            }

            this.Salt = Convert.ToBase64String(bySalt);
            this.Hash = Convert.ToBase64String(hash);            
        }
        /// <summary>
        /// The security descriptor defines user accounts who may edit this range without providing a password to access the range.
        /// </summary>
        public string SecurityDescriptor
        {
            get
            {
                return this.GetXmlNodeString("@securityDescriptor");
            }
            set
            {
                this.SetXmlNodeString("@securityDescriptor",value);
            }
        }
        internal int SpinCount
        {
            get
            {
                return this.GetXmlNodeInt("@spinCount");
            }
            set
            {
                this.SetXmlNodeString("@spinCount",value.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal string Salt
        {
            get
            {
                return this.GetXmlNodeString("@saltValue");
            }
            set
            {
                this.SetXmlNodeString("@saltValue", value);
            }
        }
        internal string Hash
        {
            get
            {
                return this.GetXmlNodeString("@hashValue");
            }
            set
            {
                this.SetXmlNodeString("@hashValue", value);
            }
        }
        internal eProtectedRangeAlgorithm Algorithm
        {
            get
            {
                string? v= this.GetXmlNodeString("@algorithmName");
                return (eProtectedRangeAlgorithm)Enum.Parse(typeof(eProtectedRangeAlgorithm), v.Replace("-", ""));
            }
            set
            {
                string? v = value.ToString();
                if(v.StartsWith("SHA"))
                {
                    v=v.Insert(3,"-");
                }
                else if(v.StartsWith("RIPEMD"))
                {
                    v=v.Insert(6,"-");
                }

                this.SetXmlNodeString("@algorithmName", v);
            }
        }
    }
}
