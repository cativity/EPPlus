﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal class FontScaleFactor
    {
        public FontScaleFactor(float small, float medium, float large)
            : this(small, medium, large, 1f)
        {

        }

        public FontScaleFactor(float small, float medium, float large, float sizeFactor)
        {
            this._small = small;
            this._medium = medium;
            this._large = large;
            this._sizeFactor = sizeFactor;
        }

        private readonly float _small;
        private readonly float _medium;
        private readonly float _large;
        private readonly float _sizeFactor;

        internal float Calculate(float width)
        {
            if (width < (100 * this._sizeFactor))
            {
                return Adjustment(width, (25 * this._sizeFactor), (100 * this._sizeFactor), this._small, this._medium);
            }
            else if (width < (200 * this._sizeFactor))
            {
                return Adjustment(width, (100 * this._sizeFactor), (200 * this._sizeFactor), this._medium, this._large);
            }
            else
            {
                return this._large;
            }
        }

        private static float Adjustment(float v, float lowerWidth, float upperWidth, float originalFactorLower, float originalFactorUpper)
        {
            if (v < lowerWidth)
            {
                return originalFactorLower;
            }

            if (v > upperWidth)
            {
                return originalFactorLower;
            }

            float f = originalFactorUpper - originalFactorLower;
            float f2 = v / upperWidth;
            return originalFactorLower + (f * f2);
        }
    }
}
