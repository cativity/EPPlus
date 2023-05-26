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

namespace OfficeOpenXml.Core.Worksheet.Fonts;

internal class BoundingRectangle
{
    public BoundingRectangle(short xMin, short yMin, short xMax, short yMax)
    {
        this.Xmin = xMin;
        this.Ymin = yMin;
        this.Xmax = xMax;
        this.Ymax = yMax;
    }

    public short Xmin { get; set; }

    public short Ymin { get; set; }

    public short Xmax { get; set; }

    public short Ymax { get; set; }
}