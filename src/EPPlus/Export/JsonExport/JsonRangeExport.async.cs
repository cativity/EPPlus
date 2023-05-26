using OfficeOpenXml.Export.HtmlExport;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    internal partial class JsonRangeExport : JsonExport
    {
        internal async Task ExportAsync(Stream stream)
        {
            StreamWriter? sw = new StreamWriter(stream);
            await this.WriteStartAsync(sw);
            await this.WriteItemAsync(sw, $"\"{this._settings.RootElementName}\":");
            await this.WriteStartAsync(sw);
            if (this._settings.FirstRowIsHeader || (this._settings.AddDataTypesOn == eDataTypeOn.OnColumn && this._range.Rows > 1))
            {
                await this.WriteColumnDataAsync(sw);
            }
            await this.WriteCellDataAsync(sw, this._range, this._settings.FirstRowIsHeader ? 1 : 0);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await this.WriteItemAsync(sw, $"\"{this._settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < this._range.Columns; i++)
            {
                await this.WriteStartAsync(sw);
                if (this._settings.FirstRowIsHeader)
                {
                    await this.WriteItemAsync(sw, $"\"name\":\"{this._range.GetCellValue<string>(0, i)}\"", false, this._settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                }
                if (this._settings.AddDataTypesOn == eDataTypeOn.OnColumn)
                {
                    string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(this._range.GetCellValue<object>(1, i));
                    await this.WriteItemAsync(sw, $"\"dt\":\"{dt}\"");
                }
                if (i == this._range.Columns - 1)
                {
                    await this.WriteEndAsync(sw, "}");
                }
                else
                {
                    await this.WriteEndAsync(sw, "},");
                }
            }

            await this.WriteEndAsync(sw, "],");
        }
    }
}
#endif