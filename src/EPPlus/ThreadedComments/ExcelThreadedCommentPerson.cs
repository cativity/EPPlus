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
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ThreadedComments
{
    /// <summary>
    /// A person in the context of ThreadedComments.
    /// Might refer to an <see cref="IdentityProvider"/>, see property ProviderId.
    /// </summary>
    public class ExcelThreadedCommentPerson : XmlHelper, IEqualityComparer<ExcelThreadedCommentPerson>
    {
        internal static string NewId()
        {
            Guid guid = Guid.NewGuid();
            return "{" + guid.ToString().ToUpper() + "}";
        }

        internal ExcelThreadedCommentPerson(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            this.TopNode = topNode;
            this.SchemaNodeOrder = new string[] { "displayName", "id", "userId", "providerId" };
        }

        /// <summary>
        /// Unique Id of the person
        /// </summary>
        public string Id
        {
            get { return this.GetXmlNodeString("@id"); }
            set { this.SetXmlNodeString("@id", value); }
        }

        /// <summary>
        /// Display name of the person
        /// </summary>
        public string DisplayName
        {
            get { return this.GetXmlNodeString("@displayName"); }
            set { this.SetXmlNodeString("@displayName", value); }
        }

        /// <summary>
        /// See the documentation of the members of the <see cref="IdentityProvider"/> enum and
        /// Microsofts documentation at https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968
        /// </summary>
        public string UserId
        {
            get { return this.GetXmlNodeString("@userId"); }
            set { this.SetXmlNodeString("@userId", value); }
        }

        /// <summary>
        /// See the documentation of the members of the <see cref="IdentityProvider"/> enum and
        /// Microsofts documentation at https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968
        /// </summary>
        public IdentityProvider ProviderId
        {
            get 
            { 
                string? id = this.GetXmlNodeString("@providerId");
                if (string.IsNullOrEmpty(this.UserId) && this.UserId == "AD")
                {
                    throw new InvalidOperationException("Cannot get ProviderId when UserId is not set");
                }

                switch(id)
                {
                    case "Windows Live":
                        return IdentityProvider.WindowsLiveId;
                    case "PeoplePicker":
                        return IdentityProvider.PeoplePicker;
                    case "AD":
                        if (this.UserId.Contains("::"))
                        {
                            return IdentityProvider.Office365;
                        }

                        return IdentityProvider.ActiveDirectory;
                    default:
                        return IdentityProvider.NoProvider;
                }
            
            }
            set 
            {
                switch(value)
                {
                    case IdentityProvider.ActiveDirectory:
                        this.SetXmlNodeString("@providerId", "AD");
                        break;
                    case IdentityProvider.WindowsLiveId:
                        this.SetXmlNodeString("@providerId", "Windows Live");
                        break;
                    case IdentityProvider.Office365:
                        this.SetXmlNodeString("@providerId", "AD");
                        break;
                    case IdentityProvider.PeoplePicker:
                        this.SetXmlNodeString("@providerId", "PeoplePicker");
                        break;
                    default:
                        this.SetXmlNodeString("@providerId", "None");
                        break;
                }
            }
        }

        /// <summary>
        /// Determines whether the specified objects are equal.
        /// </summary>
        /// <param name="x">The first object to compare.</param>
        /// <param name="y">The second object to compare.</param>
        /// <returns></returns>
        public bool Equals(ExcelThreadedCommentPerson x, ExcelThreadedCommentPerson y)
        {
            if (x == null && y == null)
            {
                return true;
            }

            if (x == null ^ y == null)
            {
                return false;
            }

            if (x.UserId == y.UserId)
            {
                return true;
            }

            return false;
        }
        /// <summary>
        /// Returns a hash code for the specified object.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object"/> for which a hash code is to be returned.</param>
        /// <returns></returns>
        public int GetHashCode(ExcelThreadedCommentPerson obj)
        {
            return obj.GetHashCode();
        }

        /// <summary>
        ///     Returns a string that represents the current object.
        /// </summary>
        /// <returns>A string that represents the current object.</returns>
        public override string ToString()
        {
            return this.DisplayName + " (id: " + this.UserId + ")";
        }
    }
}
