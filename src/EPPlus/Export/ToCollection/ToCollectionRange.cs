using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.ToCollection.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.ToCollection;

internal class ToCollectionRange
{
    internal static List<string> GetRangeHeaders(ExcelRangeBase range, string[] headers, int? headerRow)
    {
        List<string> headersList;

        if (headers == null || headers.Length == 0)
        {
            headersList = new List<string>();

            if (headerRow.HasValue == false)
            {
                return headersList;
            }

            for (int c = range._fromCol; c <= range._toCol; c++)
            {
                string? h = range.Worksheet.Cells[range._fromRow + headerRow.Value, c].Text;

                if (string.IsNullOrEmpty(h))
                {
                    throw new InvalidOperationException("Header cells cannot be empty");
                }

                if (headersList.Contains(h))
                {
                    throw new InvalidOperationException($"Header cells must be unique. Value : {h}");
                }

                headersList.Add(h);
            }
        }
        else
        {
            if (headers.Length > range.Columns)
            {
                throw new InvalidOperationException("ToCollectionOptions.Headers[] contain more items than the columns in the range.");
            }

            headersList = new List<string>(headers);
        }

        return headersList;
    }

    internal static List<T> ToCollection<T>(ExcelRangeBase range, Func<ToCollectionRow, T> setRow, ToCollectionRangeOptions options)
    {
        List<T>? ret = new List<T>();

        if (range._toRow < range._fromRow)
        {
            return null;
        }

        List<string>? headers = GetRangeHeaders(range, options.Headers, options.HeaderRow);

        List<ExcelValue>? values = new List<ExcelValue>();
        ToCollectionRow? row = new ToCollectionRow(headers, range._workbook, options.ConversionFailureStrategy);
        int startRow = options.DataStartRow ?? (options.HeaderRow ?? -1) + 1;

        for (int r = range._fromRow + startRow; r <= range._toRow; r++)
        {
            for (int c = range._fromCol; c <= range._toCol; c++)
            {
                values.Add(range.Worksheet.GetCoreValueInner(r, c));
            }

            row._cellValues = values;
            T? item = setRow(row);

            if (item != null)
            {
                ret.Add(item);
            }

            values.Clear();
        }

        return ret;
    }

    internal static List<T> ToCollection<T>(ExcelRangeBase range, ToCollectionRangeOptions options)
    {
        Type? t = typeof(T);
        List<string>? h = GetRangeHeaders(range, options.Headers, options.HeaderRow);

        if (h.Count <= 0)
        {
            throw new InvalidOperationException("No headers specified. Please set a ToCollectionOptions.HeaderRow or ToCollectionOptions.Headers[].");
        }

        List<MappedProperty>? mappings = ToCollectionAutomap.GetAutomapList<T>(h);
        List<T>? l = new List<T>();
        List<ExcelValue>? values = new List<ExcelValue>();
        int startRow = options.DataStartRow ?? (options.HeaderRow ?? -1) + 1;

        for (int r = range._fromRow + startRow; r <= range._toRow; r++)
        {
            T? item = (T)Activator.CreateInstance(t);

            foreach (MappedProperty? m in mappings)
            {
                object? v = range.Worksheet.GetValueInner(r, m.Index + range._fromCol);

                try
                {
                    m.PropertyInfo.SetValue(item, v, null);
                }
                catch (Exception ex)
                {
                    if (options.ConversionFailureStrategy == ToCollectionConversionFailureStrategy.Exception)
                    {
                        throw new EPPlusDataTypeConvertionException($"Failure to convert value {v} for index {m.Index}", ex);
                    }
                    else
                    {
                        m.PropertyInfo.SetValue(item, default(T), null);
                    }
                }
            }

            l.Add(item);
        }

        return l;
    }
}