﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class CellDataWriter
    {
        private readonly CompileResultFactory _compileResultFactory = new CompileResultFactory();
        public void Write(ExcelRangeBase cell, EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            writer.RenderBeginTag(HtmlElements.TableData);
            // TODO: apply format    
            writer.Write(cell.Value.ToString());
            writer.RenderEndTag();
            writer.ApplyFormat(options.FormatHtml);
        }
    }
}
