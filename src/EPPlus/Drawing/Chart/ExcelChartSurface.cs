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
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Chart surface settings
    /// </summary>
    public class ExcelChartSurface : XmlHelper, IDrawingStyleBase
    {
        ExcelChart _chart;
        internal ExcelChartSurface(ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           this.SchemaNodeOrder = new string[] { "thickness", "spPr", "pictureOptions" };
           this._chart = chart;
       }
       #region "Public properties"
        const string THICKNESS_PATH = "c:thickness/@val";
       /// <summary>
       /// Show the values 
       /// </summary>
        public int Thickness
       {
           get
           {
               return this.GetXmlNodeInt(THICKNESS_PATH);
           }
           set
           {
               if(value < 0 && value > 9)
               {
                   throw (new ArgumentOutOfRangeException("Thickness out of range. (0-9)"));
               }

               this.SetXmlNodeString(THICKNESS_PATH, value.ToString());
           }
       }
       ExcelDrawingFill _fill = null;
       /// <summary>
       /// Access fill properties
       /// </summary>
       public ExcelDrawingFill Fill
       {
           get
           {
               if (this._fill == null)
               {
                   this._fill = new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);
               }
               return this._fill;
           }
       }
       ExcelDrawingBorder _border = null;
       /// <summary>
       /// Access border properties
       /// </summary>
       public ExcelDrawingBorder Border
       {
           get
           {
               if (this._border == null)
               {
                   this._border = new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder);
               }
               return this._border;
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
                if (this._effect == null)
                {
                    this._effect = new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:effectLst", this.SchemaNodeOrder);
                }
                return this._effect;
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
                if (this._threeD == null)
                {
                    this._threeD = new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);
                }
                return this._threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            this.CreatespPrNode();
        }

        #endregion
    }
}
