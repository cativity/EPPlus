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
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name",value);
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
                if(_address==null)
                {
                    _address=new ExcelAddress(GetXmlNodeString("@sqref"));
                }
                return _address;
            }
            set
            {
                SetXmlNodeString("@sqref", SqRefUtility.ToSqRefAddress(value.Address));
                _address=value;
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
            Algorithm=eProtectedRangeAlgorithm.SHA512;
            SpinCount = SpinCount < 100000 ? 100000 : SpinCount;

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
            for (int i = 0; i < SpinCount; i++)
            {
                buffer=new byte[hash.Length+4];
                Array.Copy(hash, buffer, hash.Length);
                Array.Copy(BitConverter.GetBytes(i), 0, buffer, hash.Length, 4);
                hash = hp.ComputeHash(buffer);
            }
            Salt = Convert.ToBase64String(bySalt);
            Hash = Convert.ToBase64String(hash);            
        }
        /// <summary>
        /// The security descriptor defines user accounts who may edit this range without providing a password to access the range.
        /// </summary>
        public string SecurityDescriptor
        {
            get
            {
                return GetXmlNodeString("@securityDescriptor");
            }
            set
            {
                SetXmlNodeString("@securityDescriptor",value);
            }
        }
        internal int SpinCount
        {
            get
            {
                return GetXmlNodeInt("@spinCount");
            }
            set
            {
                SetXmlNodeString("@spinCount",value.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal string Salt
        {
            get
            {
                return GetXmlNodeString("@saltValue");
            }
            set
            {
                SetXmlNodeString("@saltValue", value);
            }
        }
        internal string Hash
        {
            get
            {
                return GetXmlNodeString("@hashValue");
            }
            set
            {
                SetXmlNodeString("@hashValue", value);
            }
        }
        internal eProtectedRangeAlgorithm Algorithm
        {
            get
            {
                string? v=GetXmlNodeString("@algorithmName");
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
                SetXmlNodeString("@algorithmName", v);
            }
        }
    }
}
