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

#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal abstract partial class HtmlWriterBase
    {
        public async Task WriteLineAsync()
        {
            this._newLine = true;
            await this._writer.WriteLineAsync();
        }

        public async Task WriteAsync(string text) => await this._writer.WriteAsync(text);

        internal protected async Task WriteIndentAsync()
        {
            for (int x = 0; x < this.Indent; x++)
            {
                await this._writer.WriteAsync(IndentWhiteSpace);
            }
        }

        internal async Task ApplyFormatAsync(bool minify)
        {
            if (minify == false)
            {
                await this.WriteLineAsync();
            }
        }

        internal async Task ApplyFormatIncreaseIndentAsync(bool minify)
        {
            if (minify == false)
            {
                await this.WriteLineAsync();
                this.Indent++;
            }
        }

        internal async Task ApplyFormatDecreaseIndentAsync(bool minify)
        {
            if (minify == false)
            {
                await this.WriteLineAsync();
                this.Indent--;
            }
        }

        internal async Task WriteClassAsync(string value, bool minify)
        {
            if (minify)
            {
                await this._writer.WriteAsync(value);
            }
            else
            {
                await this._writer.WriteLineAsync(value);
                this.Indent = 1;
            }
        }

        internal async Task WriteClassEndAsync(bool minify)
        {
            if (minify)
            {
                await this._writer.WriteAsync("}");
            }
            else
            {
                await this._writer.WriteLineAsync("}");
                this.Indent = 0;
            }
        }

        internal async Task WriteCssItemAsync(string value, bool minify)
        {
            if (minify)
            {
                await this._writer.WriteAsync(value);
            }
            else
            {
                await this.WriteIndentAsync();
                this._writer.WriteLine(value);
            }
        }
    }
#endif
}