using OfficeOpenXml.Export.HtmlExport;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    internal partial class JsonTableExport : JsonExport
    {
        internal async Task ExportAsync(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            await this.WriteStartAsync(sw);
            await this.WriteItemAsync(sw, $"\"{this._settings.RootElementName}\":");
            await this.WriteStartAsync(sw);

            if (this._settings.WriteNameAttribute)
            {
                await this.WriteItemAsync(sw, $"\"name\":\"{JsonEscape(this._table.Name)}\",");
            }

            if (this._settings.WriteShowHeaderAttribute)
            {
                await this.WriteItemAsync(sw, $"\"showHeader\":\"{(this._table.ShowHeader ? "1" : "0")}\",");
            }

            if (this._settings.WriteShowTotalsAttribute)
            {
                await this.WriteItemAsync(sw, $"\"showTotal\":\"{(this._table.ShowTotal ? "1" : "0")}\",");
            }

            if (this._settings.WriteColumnsElement)
            {
                await this.WriteColumnDataAsync(sw);
            }

            await this.WriteCellDataAsync(sw, this._table.DataRange, 0);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await this.WriteItemAsync(sw, $"\"{this._settings.ColumnsElementName}\":[", true);

            for (int i = 0; i < this._table.Columns.Count; i++)
            {
                await this.WriteStartAsync(sw);
                await this.WriteItemAsync(sw, $"\"name\":\"{this._table.Columns[i].Name}\"", false, this._settings.AddDataTypesOn == eDataTypeOn.OnColumn);

                if (this._settings.AddDataTypesOn == eDataTypeOn.OnColumn)
                {
                    string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(this._table.DataRange.GetCellValue<object>(0, i));
                    await this.WriteItemAsync(sw, $"\"dt\":\"{dt}\"");
                }

                if (i == this._table.Columns.Count - 1)
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