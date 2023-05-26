/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           Initial release EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Color variation for a region map chart series
    /// </summary>
    public class ExcelChartExValueColors : XmlHelper
    {
        private ExcelRegionMapChartSerie _series;

        internal ExcelChartExValueColors(ExcelRegionMapChartSerie series, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder) : base (nameSpaceManager, topNode)
        {
            this._series = series;
            this.SchemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// Number of colors to create the series gradient color scale.
        /// If two colors, the mid color is null.
        /// </summary>
        public eNumberOfColors NumberOfColors 
        { 
            get
            {
                string? v= this.GetXmlNodeString("cx:valueColorPositions/@count");
                if(v=="3")
                {
                    return eNumberOfColors.ThreeColor;
                }
                else
                {
                    return eNumberOfColors.TwoColor;
                }
            }
            set
            {
                this.SetXmlNodeString("cx:valueColorPositions/@count", ((int)value).ToString(CultureInfo.InvariantCulture));
            }
        }
        ExcelChartExValueColor _minColor = null;
        /// <summary>
        /// The minimum color value.
        /// </summary>
        public ExcelChartExValueColor MinColor 
        {
            get
            {
                if(this._minColor==null)
                {
                    this._minColor = new ExcelChartExValueColor(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, "min");
                }
                return this._minColor;
            }
        }
        ExcelChartExValueColor _midColor = null;
        /// <summary>
        /// The mid color value. Null if NumberOfcolors is set to TwoColors
        /// </summary>
        public ExcelChartExValueColor MidColor
        {
            get
            {
                if (this.NumberOfColors == eNumberOfColors.TwoColor)
                {
                    return null;
                }

                if (this._midColor == null)
                {
                    this._midColor = new ExcelChartExValueColor(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, "mid");
                }
                return this._midColor;
            }
        }
        ExcelChartExValueColor _maxColor = null;
        /// <summary>
        /// The maximum color value.
        /// </summary>
        public ExcelChartExValueColor MaxColor
        {
            get
            {
                if (this._maxColor == null)
                {
                    this._maxColor = new ExcelChartExValueColor(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, "max");
                }
                return this._maxColor;
            }
        }
    }
}