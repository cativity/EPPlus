using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml;

internal partial class JsonTableExport : JsonExport
{
    private ExcelTable _table;
    private JsonTableExportSettings _settings;

    internal JsonTableExport(ExcelTable table, JsonTableExportSettings settings)
        : base(settings)
    {
        this._table = table;
        this._settings = settings;
    }

    internal void Export(Stream stream)
    {
        StreamWriter sw = new StreamWriter(stream);
        this.WriteStart(sw);
        this.WriteItem(sw, $"\"{this._settings.RootElementName}\":");
        this.WriteStart(sw);

        if (this._settings.WriteNameAttribute)
        {
            this.WriteItem(sw, $"\"name\":\"{JsonEscape(this._table.Name)}\",");
        }

        if (this._settings.WriteShowHeaderAttribute)
        {
            this.WriteItem(sw, $"\"showHeader\":\"{(this._table.ShowHeader ? "1" : "0")}\",");
        }

        if (this._settings.WriteShowTotalsAttribute)
        {
            this.WriteItem(sw, $"\"showTotal\":\"{(this._table.ShowTotal ? "1" : "0")}\",");
        }

        if (this._settings.WriteColumnsElement)
        {
            this.WriteColumnData(sw);
        }

        this.WriteCellData(sw, this._table.DataRange, 0);
        sw.Write("}");
        sw.Flush();
    }

    private void WriteColumnData(StreamWriter sw)
    {
        this.WriteItem(sw, $"\"{this._settings.ColumnsElementName}\":[", true);

        for (int i = 0; i < this._table.Columns.Count; i++)
        {
            this.WriteStart(sw);
            this.WriteItem(sw, $"\"name\":\"{this._table.Columns[i].Name}\"", false, this._settings.AddDataTypesOn == eDataTypeOn.OnColumn);

            if (this._settings.AddDataTypesOn == eDataTypeOn.OnColumn)
            {
                string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(this._table.DataRange.GetCellValue<object>(0, i));
                this.WriteItem(sw, $"\"dt\":\"{dt}\"");
            }

            if (i == this._table.Columns.Count - 1)
            {
                this.WriteEnd(sw, "}");
            }
            else
            {
                this.WriteEnd(sw, "},");
            }
        }

        this.WriteEnd(sw, "],");
    }
}