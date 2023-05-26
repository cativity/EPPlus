/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sorting
{
    /// <summary>
    /// Represents a sort condition within a sort
    /// </summary>
    public class SortCondition : XmlHelper
    {
        internal SortCondition(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }

        private string _descendingPath = "@descending";
        private string _refPath = "@ref";
        private string _customListPath = "@customList";

        /// <summary>
        /// Sort direction of this condition. If false - ascending, if true - descending.
        /// </summary>
        public bool Descending
        {
            get
            {
                return this.GetXmlNodeBool(this._descendingPath);
            }
            set
            {
                this.SetXmlNodeBool(this._descendingPath, value);
            }
        }

        /// <summary>
        /// Address of the range used by this condition.
        /// </summary>
        public string Ref 
        {
            get
            {
                return this.GetXmlNodeString(this._refPath);
            }
            set
            {
                this.SetXmlNodeString(this._refPath, value);
            }
        }

        /// <summary>
        /// A custom list of strings that defines the sort order for this condition.
        /// </summary>
        public string[] CustomList
        {
            get
            {
                string? list = this.GetXmlNodeString(this._customListPath);
                if(!string.IsNullOrEmpty(list))
                {
                    return list.Split(',').Where(x => !string.IsNullOrEmpty(x)).Select(x => x.Trim()).ToArray();
                }
                return null;
            }
            set
            {
                if(value == null || value.Length == 0)
                {
                    this.SetXmlNodeString(this._customListPath, string.Empty, true);
                }
                StringBuilder? val = new StringBuilder();
                for(int x = 0; x < value.Length; x++)
                {
                    val.Append(value[x]);
                    if(x < value.Length -1)
                    {
                        val.Append(",");
                    }
                }

                this.SetXmlNodeString(this._customListPath, val.ToString());
            }
        }
    }
}
