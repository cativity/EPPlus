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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents a marker on a chart serie
    /// </summary>
    public class ExcelChartMarker : XmlHelper, IDrawingStyleBase
    {
        ExcelChart _chart;
        bool _allowMarkers;
        internal ExcelChartMarker(ExcelChart chart,XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "symbol", "size", "spPr"});
            this._chart = chart;
            this._allowMarkers = chart.IsType3D();
        }
        /// <summary>
        /// The marker style
        /// </summary>
        public eMarkerStyle Style
        {
            get
            {
                return this.GetXmlNodeString("c:marker/c:symbol/@val").ToEnum(eMarkerStyle.None);
            }
            set
            {
                if(this._allowMarkers)
                {
                    throw (new ArgumentException("Style", "Can't set markers on a 3d chart serie"));
                }

                this.SetXmlNodeString("c:marker/c:symbol/@val", value.ToEnumString());
            }
        }
        /// <summary>
        /// The size of the marker.
        /// Ranges from 2 to 72.
        /// </summary>
        public int Size
        {
            get
            {
                int v= this.GetXmlNodeInt("c:marker/c:size/@val");
                if(v<0)
                {
                    return 5;   //Default value;
                }
                return v;
            }
            set
            {
                if (this._allowMarkers)
                {
                    throw (new ArgumentException("Size", "Can't set markers on a 3d chart serie"));
                }

                if (value<2 || value>72)
                {
                    throw (new ArgumentOutOfRangeException("Marker size must be between 2 and 72"));
                }

                this.SetXmlNodeString("c:marker/c:size/@val", value.ToString(CultureInfo.InvariantCulture));
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (this._allowMarkers)
                {
                    throw (new ArgumentException("Fill", "Can't set markers on a 3d chart serie"));
                }

                return this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:marker/c:spPr", this.SchemaNodeOrder);
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// A reference to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (this._allowMarkers)
                {
                    throw (new ArgumentException("Border", "Can't set markers on a 3d chart serie"));
                }

                return this._border ??= new ExcelDrawingBorder(this._chart,
                                                               this.NameSpaceManager,
                                                               this.TopNode,
                                                               "c:marker/c:spPr/a:ln",
                                                               this.SchemaNodeOrder);
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (this._allowMarkers)
                {
                    throw (new ArgumentException("Effect", "Can't set markers on a 3d chart serie"));
                }

                return this._effect ??= new ExcelDrawingEffectStyle(this._chart,
                                                                    this.NameSpaceManager,
                                                                    this.TopNode,
                                                                    "c:marker/c:spPr/a:effectLst",
                                                                    this.SchemaNodeOrder);
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (this._allowMarkers)
                {
                    throw (new ArgumentException("ThreeD", "Can't set markers on a 3d chart serie"));
                }

                return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:marker/c:spPr", this.SchemaNodeOrder);
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            this.CreatespPrNode();
        }

    }
}
