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
    using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Sparkline colors
    /// </summary>
    public class ExcelSparklineColor : XmlHelper, IColor
    {
        internal ExcelSparklineColor(XmlNamespaceManager ns, XmlNode node) : base(ns, node)
        {

        }
        /// <summary>
        /// Indexed color
        /// </summary>
        public int Indexed
        {
            get => this.GetXmlNodeInt("@indexed");
            set
            {
                if (value < 0 || value > 65)
                {
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }

                this.ClearValues();
                this.SetXmlNodeString("@indexed", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// RGB 
        /// </summary>
        public string Rgb
        {
            get => this.GetXmlNodeString("@rgb");
            internal set
            {
                this.ClearValues();
                this.SetXmlNodeString("@rgb", value);
            }
        }
        /// <summary>
        /// The theme color
        /// </summary>
        public eThemeSchemeColor? Theme 
        {
            get
            {
                int? v = this.GetXmlNodeIntNull("@theme");
                if(v.HasValue)
                {
                    return (eThemeSchemeColor)v;
                }
                else
                {
                    return null;
                }
            }
            internal set
            {
                this.ClearValues();

                this.SetXmlNodeString("@theme", ((int)value.Value).ToString(CultureInfo.InvariantCulture));
            }
        }

        private void ClearValues()
        {
            this.DeleteNode("@rgb");
            this.DeleteNode("@indexed");
            this.DeleteNode("@theme");
            this.DeleteNode("@auto");
        }

        /// <summary>
        /// The tint value
        /// </summary>
        public decimal Tint
        {
            get=> this.GetXmlNodeDecimal("@tint");
            set
            {
                if (value > 1 || value < -1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between -1 and 1"));
                }

                this.SetXmlNodeString("@tint", value.ToString(CultureInfo.InvariantCulture));
            }
        }
        /// <summary>
        /// Color is set to automatic
        /// </summary>
        public bool Auto
        {
            get
            {
                return this.GetXmlNodeBool("@auto");
            }
            internal set
            {
                this.ClearValues();
                this.SetXmlNodeBool("@auto", value);
            }
        }
        /// <summary>
        /// Sets a color
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(Color color)
        {
            this.Rgb = color.ToArgb().ToString("X");
        }
        /// <summary>
        /// Sets a theme color
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(eThemeSchemeColor color)
        {
            this.Theme=color;
        }
        /// <summary>
        /// Sets an indexed color
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(ExcelIndexedColor color)
        {
            this.Indexed = (int)color;
        }
        /// <summary>
        /// Sets the color to auto
        /// </summary>
        public void SetAuto()
        {
            this.Auto = true;
        }
    }
}
