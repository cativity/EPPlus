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

namespace OfficeOpenXml
{
    /// <summary>
    /// Error ignore options for a worksheet
    /// </summary>
    public class ExcelIgnoredError : XmlHelper
    {
        internal ExcelIgnoredError(XmlNamespaceManager nsm, XmlNode topNode, ExcelAddressBase address) : base(nsm, topNode)
        {
            this.SetXmlNodeString("@sqref", address.AddressSpaceSeparated);
        }
        /// <summary>
        /// Ignore errors when numbers are formatted as text or are preceded by an apostrophe
        /// </summary>
        public bool NumberStoredAsText
        {
            get
            {
                return this.GetXmlNodeBool("@numberStoredAsText");
            }
            set
            {
                this.SetXmlNodeBool("@numberStoredAsText", value);
            }
        }
        /// <summary>
        /// Calculated Column
        /// </summary>
        public bool CalculatedColumm
        {
            get
            {
                return this.GetXmlNodeBool("@calculatedColumn");
            }
            set
            {
                this.SetXmlNodeBool("@calculatedColumn", value);
            }
        }


        /// <summary>
        /// Ignore errors when a formula refers an empty cell
        /// </summary>
        public bool EmptyCellReference
        {
            get
            {
                return this.GetXmlNodeBool("@emptyCellReference");
            }
            set
            {
                this.SetXmlNodeBool("@emptyCellReference", value);
            }
        }

        /// <summary>
        /// Ignore errors when formulas fail to Evaluate
        /// </summary>
        public bool EvaluationError
        {
            get
            {
                return this.GetXmlNodeBool("@evalError");
            }
            set
            {
                this.SetXmlNodeBool("@evalError", value);
            }
        }
        /// <summary>
        /// Ignore errors when a formula in a region of your worksheet differs from other formulas in the same region.
        /// </summary>
        public bool Formula
        {
            get
            {
                return this.GetXmlNodeBool("@formula");
            }
            set
            {
                this.SetXmlNodeBool("@formula", value);
            }
        }
        /// <summary>
        /// Ignore errors when formulas omit certain cells in a region.
        /// </summary>
        public bool FormulaRange
        {
            get
            {
                return this.GetXmlNodeBool("@formulaRange");
            }
            set
            {
                this.SetXmlNodeBool("@formulaRange", value);
            }
        }
        /// <summary>
        /// Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified
        /// </summary>
        public bool ListDataValidation
        {
            get
            {
                return this.GetXmlNodeBool("@listDataValidation");
            }
            set
            {
                this.SetXmlNodeBool("@listDataValidation", value);
            }
        }
        /// <summary>
        /// The address
        /// </summary>
        public ExcelAddressBase Address
        {
            get
            {
                return new ExcelAddressBase(this.GetXmlNodeString("@sqref"));
            }
        }
        /// <summary>
        /// Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        /// </summary>
        public bool TwoDigitTextYear
        {
            get
            {
                return this.GetXmlNodeBool("@twoDigitTextYear");
            }
            set
            {
                this.SetXmlNodeBool("@twoDigitTextYear", value);
            }
        }
        /// <summary>
        /// Ignore errors when unlocked cells contain formulas
        /// </summary>
        public bool UnlockedFormula
        {
            get
            {
                return this.GetXmlNodeBool("@unlockedFormula");
            }
            set
            {
                this.SetXmlNodeBool("@unlockedFormula", value);
            }
        }
    }
}
