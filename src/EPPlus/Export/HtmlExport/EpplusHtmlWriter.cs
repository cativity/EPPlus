/*************************************************************************************************
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
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style.XmlAccess;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter : HtmlWriterBase
    {
        internal EpplusHtmlWriter(Stream stream, Encoding encoding, Dictionary<string, int> styleCache) : base(stream, encoding, styleCache)
        {
        }

        private readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            this._attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }
        public void RenderBeginTag(string elementName, bool closeElement = false)
        {
            this._newLine = false;
            // avoid writing indent characters for a hyperlinks or images inside a td element
            if(elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                this.WriteIndent();
            }

            this._writer.Write($"<{elementName}");
            foreach (EpplusHtmlAttribute? attribute in this._attributes)
            {
                this._writer.Write($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }

            this._attributes.Clear();

            if (closeElement)
            {
                this._writer.Write("/>");
                this._writer.Flush();
            }
            else
            {
                this._writer.Write(">");
                this._elementStack.Push(elementName);
            }
        }

        public void RenderEndTag()
        {
            if (this._newLine)
            {
                this.WriteIndent();
            }

            string? elementName = this._elementStack.Pop();
            this._writer.Write($"</{elementName}>");
            this._writer.Flush();
        }

        internal void SetClassAttributeFromStyle(ExcelRangeBase cell, bool isHeader, HtmlExportSettings settings, string additionalClasses)
        {            
            string cls = string.IsNullOrEmpty(additionalClasses) ? "" : additionalClasses;
            int styleId = cell.StyleID;
            ExcelStyles styles = cell.Worksheet.Workbook.Styles;
            if (styleId < 0 || styleId >= styles.CellXfs.Count)
            {
                return;
            }

            ExcelXfs? xfs = styles.CellXfs[styleId];
            string? styleClassPrefix = settings.StyleClassPrefix;
            if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.CellDataType &&
               xfs.HorizontalAlignment == ExcelHorizontalAlignment.General)
            {
                if (ConvertUtil.IsNumericOrDate(cell.Value))
                {
                    cls = $"{styleClassPrefix}ar";
                }
                else if (isHeader)
                {
                    cls = $"{styleClassPrefix}al";
                }
            }

            if (styleId == 0 || HasStyle(xfs) == false)
            {
                if (string.IsNullOrEmpty(cls) == false)
                {
                    this.AddAttribute("class", cls);
                }

                return;
            }

            string key = GetStyleKey(xfs);

            string? ma = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
            if (ma != null)
            {
                ExcelAddressBase? address = new ExcelAddressBase(ma);
                int bottomStyleId = cell.Worksheet._values.GetValue(address._toRow, address._fromCol)._styleId;
                int rightStyleId = cell.Worksheet._values.GetValue(address._fromRow, address._toCol)._styleId;
                key += bottomStyleId + "|" + rightStyleId;
            }

            int id;
            if (this._styleCache.ContainsKey(key))
            {
                id = this._styleCache[key];
            }
            else
            {
                id = this._styleCache.Count + 1;
                this._styleCache.Add(key, id);
            }
            cls += $" {styleClassPrefix}{settings.CellStyleClassName}{id}";
            this.AddAttribute("class", cls.Trim());
        }
    }
}
