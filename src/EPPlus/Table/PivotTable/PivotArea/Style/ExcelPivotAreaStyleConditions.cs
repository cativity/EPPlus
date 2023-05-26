/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/18/2021         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Conditions for a pivot table area style.
    /// </summary>
    public class ExcelPivotAreaStyleConditions
    {
        internal ExcelPivotAreaStyleConditions(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt)
        {
            this.Fields = new ExcelPivotAreaReferenceCollection(nsm, topNode, pt);
            XmlHelper? xh = XmlHelperFactory.Create(nsm, topNode);
            foreach (XmlElement n in xh.GetNodes("d:references/d:reference"))
            {
                if (n.GetAttribute("field") == "4294967294")
                {
                    this.DataFields = new ExcelPivotAreaDataFieldReference(nsm, n, pt, -2);
                }
                else
                {
                    this.Fields.Add(new ExcelPivotAreaReference(nsm, n, pt));
                }
            }

            this.DataFields ??= new ExcelPivotAreaDataFieldReference(nsm, topNode, pt, -2);
        }
        /// <summary>
        /// Row and column fields that the conditions will apply to. 
        /// </summary>
        public ExcelPivotAreaReferenceCollection Fields 
        { 
            get;  
        }
        /// <summary>
        /// The data field that the conditions will apply to. 
        /// </summary>
        public ExcelPivotAreaDataFieldReference DataFields
        {
            get;
        }

        internal void UpdateXml()
        {
            this.DataFields.UpdateXml();
            foreach (ExcelPivotAreaReference r in this.Fields)
            {
                r.UpdateXml();
            }
        }
    }
}
