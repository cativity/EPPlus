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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Information about an address.
    /// </summary>
    public class ExcelAddressInfo
    {
        private ExcelAddressInfo(string address) 
        {   
            string? addressOnSheet = address;
            this.Worksheet = string.Empty;
            if (address.Contains("!"))
            {
                string[]? worksheetArr = address.Split('!');
                this.Worksheet = worksheetArr[0];
                addressOnSheet = worksheetArr[1];
            }
            if (addressOnSheet.Contains(":"))
            {
                string[]? rangeArr = addressOnSheet.Split(':');
                this.StartCell = rangeArr[0];
                this.EndCell = rangeArr[1];
            }
            else
            {
                this.StartCell = addressOnSheet;
            }

            this.AddressOnSheet = addressOnSheet;
        }

        public static ExcelAddressInfo Parse(string address)
        {
            Require.That(address).Named("address").IsNotNullOrEmpty();
            return new ExcelAddressInfo(address);
        }

        /// <summary>
        /// The worksheet name
        /// </summary>
        public string Worksheet { get; private set; }

        /// <summary>
        /// Returns true if the <see cref="Worksheet"/> is set
        /// </summary>
        public bool WorksheetIsSpecified
        {
            get
            {
                return !string.IsNullOrEmpty(this.Worksheet);
            }
        }

        /// <summary>
        /// If the address reference multiple cells.
        /// </summary>
        public bool IsMultipleCells 
        { 
            get 
            { 
                return !string.IsNullOrEmpty(this.EndCell); 
            } 
        }

        /// <summary>
        /// The start cell address
        /// </summary>
        public string StartCell { get; private set; }

        /// <summary>
        /// The end cell address
        /// </summary>
        public string EndCell { get; private set; }

        /// <summary>
        /// The address part if a worksheet is specified on the address. 
        /// </summary>
        public string AddressOnSheet { get; private set; }
    }
}
