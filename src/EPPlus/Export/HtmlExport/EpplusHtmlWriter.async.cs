/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter
    {
#if !NET35 && !NET40
        public async Task RenderBeginTagAsync(string elementName, bool closeElement = false)
        {
            this._newLine = false;

            if (elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                await this.WriteIndentAsync();
            }

            await this._writer.WriteAsync($"<{elementName}");

            foreach (EpplusHtmlAttribute? attribute in this._attributes)
            {
                await this._writer.WriteAsync($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }

            this._attributes.Clear();

            if (closeElement)
            {
                await this._writer.WriteAsync("/>");
                await this._writer.FlushAsync();
            }
            else
            {
                await this._writer.WriteAsync(">");
                this._elementStack.Push(elementName);
            }
        }

        public async Task RenderEndTagAsync()
        {
            if (this._newLine)
            {
                await this.WriteIndentAsync();
            }

            string? elementName = this._elementStack.Pop();
            await this._writer.WriteAsync($"</{elementName}>");
            await this._writer.FlushAsync();
        }
#endif
    }
}