using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml;

internal partial class JsonRangeExport : JsonExport
{
    private ExcelRangeBase _range;
    private JsonRangeExportSettings _settings;
    internal JsonRangeExport(ExcelRangeBase range, JsonRangeExportSettings settings) : base(settings)
    {
        this._range = range;
        this._settings = settings;
    }
    internal void Export(Stream stream)
    {
        StreamWriter? sw = new StreamWriter(stream);
        this.WriteStart(sw);
        this.WriteItem(sw, $"\"{this._settings.RootElementName}\":");
        this.WriteStart(sw);
        if (this._settings.FirstRowIsHeader || (this._settings.AddDataTypesOn==eDataTypeOn.OnColumn && this._range.Rows>1))
        {
            this.WriteColumnData(sw);
        }

        this.WriteCellData(sw, this._range, this._settings.FirstRowIsHeader ? 1 : 0);
        sw.Write("}");
        sw.Flush();
    }

    private void WriteColumnData(StreamWriter sw)
    {
        this.WriteItem(sw, $"\"{this._settings.ColumnsElementName}\":[", true);
        for (int i = 0; i < this._range.Columns; i++)
        {
            //if (i > 0) sw.Write(",");
            //sw.Write("{");
            this.WriteStart(sw);
            if (this._settings.FirstRowIsHeader)
            {
                this.WriteItem(sw, $"\"name\":\"{this._range.GetCellValue<string>(0,i)}\"", false, this._settings.AddDataTypesOn == eDataTypeOn.OnColumn);
            }
            if (this._settings.AddDataTypesOn==eDataTypeOn.OnColumn)
            {
                string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(this._range.GetCellValue<object>(1, i));
                this.WriteItem(sw, $"\"dt\":\"{dt}\"");
            }
            if (i == this._range.Columns - 1)
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