/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  06/05/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Slicer;
using System;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides a simple way to type cast drawing object top its top level class.
    /// </summary>
    public class ExcelDrawingAsType
    {
        ExcelDrawing _drawing;
        internal ExcelDrawingAsType(ExcelDrawing drawing)
        {
            this._drawing = drawing;
        }
        /// <summary>
        /// Converts the drawing to it's top level or other nested drawing class.        
        /// </summary>
        /// <typeparam name="T">The type of drawing. T must be inherited from ExcelDrawing</typeparam>
        /// <returns>The drawing as type T</returns>
        public T Type<T>() where T : ExcelDrawing
        {
            return this._drawing as T;
        }
        /// <summary>
        /// Returns the drawing as a shape. 
        /// If this drawing is not a shape, null will be returned
        /// </summary>
        /// <returns>The drawing as a shape</returns>
        public ExcelShape Shape
        {
            get
            {
                return this._drawing as ExcelShape;
            }
        }
        /// <summary>
        /// Returns the drawing as a picture/image. 
        /// If this drawing is not a picture, null will be returned
        /// </summary>
        /// <returns>The drawing as a picture</returns>
        public ExcelPicture Picture
        {
            get
            {
                return this._drawing as ExcelPicture;
            }
        }
        ExcelChartAsType _chartAsType;
        /// <summary>
        /// An object that containing properties that type-casts the drawing to a chart.
        /// </summary>
        public ExcelChartAsType Chart
        {
            get
            {
                if (this._chartAsType == null)
                {
                    this._chartAsType = new ExcelChartAsType(this._drawing);
                }
                return this._chartAsType;
            }
        }

        ExcelSlicerAsType _slicerAsType;
        /// <summary>
        /// An object that containing properties that type-casts the drawing to a slicer.
        /// </summary>
        public ExcelSlicerAsType Slicer 
        { 
            get
            {
                if (this._slicerAsType == null)
                {
                    this._slicerAsType = new ExcelSlicerAsType(this._drawing);
                }
                return this._slicerAsType;
            }
        }

        ExcelControlAsType _controlAsType;

        /// <summary>
        /// Helps to cast drawings to controls. Use the properties of this class to cast to the various specific control types.
        /// </summary>
        /// <returns></returns>
        public ExcelControlAsType Control
        {
            get
            {
                if(this._controlAsType == null)
                {
                    this._controlAsType = new ExcelControlAsType(this._drawing);
                }
                return this._controlAsType;
            }
        }
    }
}
