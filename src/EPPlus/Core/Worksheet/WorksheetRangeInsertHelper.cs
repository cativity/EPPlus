﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  02/03/2020         EPPlus Software AB       Added
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Sparkline;

namespace OfficeOpenXml.Core.Worksheet;

internal static class WorksheetRangeInsertHelper
{
    internal static void InsertRow(ExcelWorksheet ws, int rowFrom, int rows, int copyStylesFromRow)
    {
        ValidateInsertRow(ws, rowFrom, rows);

        lock (ws)
        {
            InsertCellStores(ws, rowFrom, 0, rows, 0);

            FixFormulasInsertRow(ws, rowFrom, rows);

            WorksheetRangeHelper.FixMergedCellsRow(ws, rowFrom, rows, false);

            if (copyStylesFromRow > 0)
            {
                CopyFromStyleRow(ws, rowFrom, rows, copyStylesFromRow);
            }

            InsertRowTable(ws, rowFrom, rows);
            InsertRowPivotTable(ws, rowFrom, rows);

            ExcelRange? range = ws.Cells[rowFrom, 1, rowFrom + rows - 1, ExcelPackage.MaxColumns];
            ExcelAddressBase? affectedAddress = GetAffectedRange(range, eShiftTypeInsert.Down);
            InsertFilterAddress(range, affectedAddress, eShiftTypeInsert.Down);
            InsertSparkLinesAddress(range, eShiftTypeInsert.Down, affectedAddress);
            InsertDataValidation(range, eShiftTypeInsert.Down, affectedAddress, ws, false);
            InsertConditionalFormatting(range, eShiftTypeInsert.Down, affectedAddress, ws, false);

            WorksheetRangeCommonHelper.AdjustDvAndCfFormulasRow(ws, rowFrom, rows);

            WorksheetRangeHelper.AdjustDrawingsRow(ws, rowFrom, rows);
        }
    }

    private static void InsertRowTable(ExcelWorksheet ws, int rowFrom, int rows)
    {
        foreach (ExcelTable? tbl in ws.Tables)
        {
            tbl.Address = tbl.Address.AddRow(rowFrom, rows);
            foreach (ExcelTableColumn? col in tbl.Columns)
            {
                if (string.IsNullOrEmpty(col.CalculatedColumnFormula) == false)
                {
                    col.CalculatedColumnFormula = ExcelCellBase.UpdateFormulaReferences(col.CalculatedColumnFormula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                }
            }
        }
    }

    private static void InsertRowPivotTable(ExcelWorksheet ws, int rowFrom, int rows)
    {
        foreach (ExcelPivotTable? ptbl in ws.PivotTables)
        {
            ptbl.Address = ptbl.Address.AddRow(rowFrom, rows);
            ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddRow(rowFrom, rows).Address;
        }
    }

    internal static void InsertColumn(ExcelWorksheet ws, int columnFrom, int columns, int copyStylesFromColumn)
    {
        ValidateInsertColumn(ws, columnFrom, columns);

        lock (ws)
        {
            InsertCellStores(ws, 0, columnFrom, 0, columns);

            FixFormulasInsertColumn(ws, columnFrom, columns);

            WorksheetRangeHelper.FixMergedCellsColumn(ws, columnFrom, columns, false);

            AdjustColumns(ws, columnFrom, columns);

            CopyStylesFromColumn(ws, columnFrom, columns, copyStylesFromColumn);

            InsertColumnTable(ws, columnFrom, columns);
            InsertColumnPivotTable(ws, columnFrom, columns);

            ExcelRange? range = ws.Cells[1, columnFrom, ExcelPackage.MaxRows, columnFrom + columns - 1];
            ExcelAddressBase? affectedAddress = GetAffectedRange(range, eShiftTypeInsert.Right);
            InsertFilterAddress(range, affectedAddress, eShiftTypeInsert.Right);
            InsertSparkLinesAddress(range, eShiftTypeInsert.Right, affectedAddress);
            InsertDataValidation(range, eShiftTypeInsert.Right, affectedAddress, ws, false);
            InsertConditionalFormatting(range, eShiftTypeInsert.Right, affectedAddress, ws, false);

            WorksheetRangeCommonHelper.AdjustDvAndCfFormulasColumn(ws, columnFrom, columns);

            //Adjust drawing positions.
            WorksheetRangeHelper.AdjustDrawingsColumn(ws, columnFrom, columns);
        }
    }

    private static void InsertColumnPivotTable(ExcelWorksheet ws, int columnFrom, int columns)
    {
        foreach (ExcelPivotTable? ptbl in ws.PivotTables)
        {
            if (columnFrom <= ptbl.Address.End.Column)
            {
                ptbl.Address = ptbl.Address.AddColumn(columnFrom, columns);
            }
            if (columnFrom <= ptbl.CacheDefinition.SourceRange.End.Column)
            {
                if (ptbl.CacheDefinition.CacheSource == eSourceType.Worksheet)
                {
                    ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddColumn(columnFrom, columns).Address;
                }
            }
        }
    }

    private static void InsertColumnTable(ExcelWorksheet ws, int columnFrom, int columns)
    {
        //Adjust tables
        foreach (ExcelTable? tbl in ws.Tables)
        {
            if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
            {
                InsertTableColumns(columnFrom, columns, tbl);
            }

            tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
            if (columnFrom <= tbl.Address._toCol)
            {
                foreach (ExcelTableColumn? col in tbl.Columns)
                {
                    if (string.IsNullOrEmpty(col.CalculatedColumnFormula) == false)
                    {
                        col.CalculatedColumnFormula = ExcelCellBase.UpdateFormulaReferences(col.CalculatedColumnFormula, 0, columns, 0, columnFrom, ws.Name, ws.Name);
                    }
                }
            }
        }
    }

    internal static void Insert(ExcelRangeBase range, eShiftTypeInsert shift, bool styleCopy, bool isTable)
    {
        ValidateInsert(range, shift);

        ExcelAddressBase? effectedAddress = GetAffectedRange(range, shift);
        WorksheetRangeHelper.ValidateIfInsertDeleteIsPossible(range, effectedAddress, GetAffectedRange(range, shift, 1), true);

        ExcelWorksheet? ws = range.Worksheet;
        lock (ws)
        {
            List<int>? styleList = GetStylesForRange(range, shift);
            WorksheetRangeHelper.ConvertEffectedSharedFormulasToCellFormulas(ws, effectedAddress);

            if (shift == eShiftTypeInsert.Down)
            {
                InsertCellStores(range._worksheet, range._fromRow, range._fromCol, range.Rows, range.Columns, range._toCol);
            }
            else
            {
                InsertCellStoreShiftRight(range._worksheet, range);
            }
            AdjustFormulasInsert(range, effectedAddress, shift);
            InsertFilterAddress(range, effectedAddress, shift);
            WorksheetRangeHelper.FixMergedCells(ws, range, shift);

            if (styleCopy)
            {
                SetStylesForRange(range, shift, styleList);
            }

            InsertTableAddress(ws, range, shift, effectedAddress);
            InsertPivottableAddress(ws, range, shift, effectedAddress);

            InsertDataValidation(range, shift, effectedAddress, ws, isTable);
            InsertConditionalFormatting(range, shift, effectedAddress, ws, isTable);

            InsertSparkLinesAddress(range, shift, effectedAddress);

            if (shift == eShiftTypeInsert.Down)
            {
                WorksheetRangeHelper.AdjustDrawingsRow(ws, range._fromRow, range.Rows, range._fromCol, range._toCol);
            }
            else
            {
                WorksheetRangeHelper.AdjustDrawingsColumn(ws, range._fromCol, range.Columns, range._fromRow, range._toRow);
            }
        }
    }

    private static void InsertConditionalFormatting(ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress, ExcelWorksheet ws, bool isTable)
    {
        List<IExcelConditionalFormattingRule>? delCF = new List<IExcelConditionalFormattingRule>();
        //Update Conditional formatting references
        foreach (IExcelConditionalFormattingRule? cf in ws.ConditionalFormatting)
        {
            ExcelAddressBase? newAddress = InsertSplitAddress(cf.Address, range, effectedAddress, shift, isTable);
            if (newAddress == null)
            {
                delCF.Add(cf);
            }
            else
            {
                ExcelConditionalFormattingRule? cfr = (ExcelConditionalFormattingRule)cf;
                if (cfr.Address.Address != newAddress.Address)
                {
                    if (cfr.Address.FirstCellAddressRelative != newAddress.FirstCellAddressRelative)
                    {
                        cfr.Formula = WorksheetRangeHelper.AdjustStartCellForFormula(cfr.Formula, cfr.Address, newAddress);
                        cfr.Formula2 = WorksheetRangeHelper.AdjustStartCellForFormula(cfr.Formula2, cfr.Address, newAddress);
                    }
                        
                    cfr.Address = new ExcelAddress(newAddress.Address);
                }
            }
        }

        foreach (IExcelConditionalFormattingRule? cf in delCF)
        {
            ws.ConditionalFormatting.Remove(cf);
        }
    }

    private static void InsertDataValidation(ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress, ExcelWorksheet ws, bool isTable)
    {
        List<ExcelDataValidation>? delDV = new List<ExcelDataValidation>();
        //Update data validation references
        foreach (ExcelDataValidation dv in ws.DataValidations)
        {
            ExcelAddressBase? newAddress = InsertSplitAddress(dv.Address, range, effectedAddress, shift, isTable);
            if (newAddress == null)
            {
                delDV.Add(dv);
            }
            else
            {
                if (dv.Address.Address != newAddress.Address)
                {
                    if (dv is ExcelDataValidationWithFormula<IExcelDataValidationFormula> dvFormula)
                    {
                        if (dv.Address.FirstCellAddressRelative != newAddress.FirstCellAddressRelative)
                        {
                            dvFormula.Formula.ExcelFormula = WorksheetRangeHelper.AdjustStartCellForFormula(dvFormula.Formula.ExcelFormula, dv.Address, newAddress);
                        }
                    }
                    dv.SetAddress(newAddress.Address);
                }
                    
            }
            ws.DataValidations.InsertRangeDictionary(range, shift == eShiftTypeInsert.Right || shift == eShiftTypeInsert.EntireColumn);
        }
        foreach (ExcelDataValidation? dv in delDV)
        {
            ws.DataValidations.Remove(dv);
        }
    }

    private static void InsertFilterAddress(ExcelRangeBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift)
    {
        ExcelWorksheet? ws = range.Worksheet;
        if (ws.AutoFilterAddress != null && effectedAddress.Collide(ws.AutoFilterAddress) != ExcelAddressBase.eAddressCollition.No)
        {
            if (shift == eShiftTypeInsert.Down)
            {
                ws.AutoFilterAddress = ws.AutoFilterAddress.AddRow(range._fromRow, range.Rows);
            }
            else
            {
                ws.AutoFilterAddress = ws.AutoFilterAddress.AddColumn(range._fromCol, range.Columns);
            }
        }
    }
    private static void InsertSparkLinesAddress(ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress)
    {
        foreach (ExcelSparklineGroup? slg in range.Worksheet.SparklineGroups)
        {
            if (slg.DateAxisRange != null && effectedAddress.Collide(slg.DateAxisRange) >= ExcelAddressBase.eAddressCollition.Inside)
            {
                string address;
                if (shift == eShiftTypeInsert.Down)
                {
                    address = slg.DateAxisRange.AddRow(range._fromRow, range.Rows).Address;
                }
                else
                {
                    address = slg.DateAxisRange.AddColumn(range._fromCol, range.Columns).Address;
                }
                slg.DateAxisRange = range.Worksheet.Cells[address];
            }

            foreach (ExcelSparkline? sl in slg.Sparklines)
            {
                if (shift == eShiftTypeInsert.Down)
                {
                    if (effectedAddress.Collide(sl.RangeAddress) >= ExcelAddressBase.eAddressCollition.Inside ||
                        range.CollideFullRow(sl.RangeAddress._fromRow, sl.RangeAddress._toRow))
                    {
                        sl.RangeAddress = sl.RangeAddress.AddRow(range._fromRow, range.Rows);
                    }

                    if (sl.Cell.Row >= range._fromRow && sl.Cell.Column >= range._fromCol && sl.Cell.Column <= range._toCol)
                    {
                        sl.Cell = new ExcelCellAddress(sl.Cell.Row + range.Rows, sl.Cell.Column);
                    }
                }
                else
                {
                    if (effectedAddress.Collide(sl.RangeAddress) >= ExcelAddressBase.eAddressCollition.Inside ||
                        range.CollideFullColumn(sl.RangeAddress._fromCol, sl.RangeAddress._toCol))
                    {
                        sl.RangeAddress = sl.RangeAddress.AddColumn(range._fromCol, range.Columns);
                    }

                    if (sl.Cell.Column >= range._fromCol && sl.Cell.Row >= range._fromRow && sl.Cell.Row <= range._toRow)
                    {
                        sl.Cell = new ExcelCellAddress(sl.Cell.Row, sl.Cell.Column + range.Columns);
                    }
                }
            }
        }
    }

    private static void ValidateInsert(ExcelRangeBase range, eShiftTypeInsert shift)
    {
        if (range == null || (range.Addresses != null && range.Addresses.Count > 1))
        {
            throw new ArgumentException("Can't insert into range. ´range´ can't be null or have multiple addresses.", "range");
        }
    }

    private static ExcelAddressBase InsertSplitAddress(ExcelAddressBase address, ExcelAddressBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift, bool isTable)
    {
        if (address.Addresses == null)
        {
            return InsertSplitIndividualAddress(address, range, effectedAddress, shift, isTable);
        }
        else
        {
            string? newAddress = "";
            foreach (ExcelAddressBase? a in address.Addresses)
            {
                newAddress += InsertSplitIndividualAddress(a, range, effectedAddress, shift, isTable) + ",";
            }
            return new ExcelAddressBase(newAddress.Substring(0, newAddress.Length - 1));
        }

    }

    private static ExcelAddressBase InsertSplitIndividualAddress(ExcelAddressBase address, ExcelAddressBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift, bool isTable)
    {
        if (address.CollideFullColumn(range._fromCol, range._toCol) && (shift == eShiftTypeInsert.Down || shift == eShiftTypeInsert.EntireRow))
        {
            return address.AddRow(range._fromRow, range.Rows, false, false, isTable);
        }
        else if (address.CollideFullRow(range._fromRow, range._toRow) && (shift == eShiftTypeInsert.Right || shift == eShiftTypeInsert.EntireColumn))
        {
            return address.AddColumn(range._fromCol, range.Columns, false, false);
        }
        else
        {
            ExcelAddressBase.eAddressCollition collide = effectedAddress.Collide(address);
            if (collide == ExcelAddressBase.eAddressCollition.Partly)
            {
                ExcelAddressBase? addressToShift = effectedAddress.Intersect(address);
                ExcelAddressBase? shiftedAddress = ShiftAddress(addressToShift, range, shift);
                string? newAddress = "";
                if (address._fromRow < addressToShift._fromRow)
                {
                    newAddress = ExcelCellBase.GetAddress(address._fromRow, address._fromCol, addressToShift._fromRow - 1, address._toCol) + ",";
                }
                if (address._fromCol < addressToShift._fromCol)
                {
                    int fromRow = Math.Max(address._fromRow, addressToShift._fromRow);
                    newAddress += ExcelCellBase.GetAddress(fromRow, address._fromCol, address._toRow, addressToShift._fromCol - 1) + ",";
                }

                newAddress += $"{shiftedAddress},";

                if (address._toRow > addressToShift._toRow)
                {
                    newAddress += ExcelCellBase.GetAddress(addressToShift._toRow + 1, address._fromCol, address._toRow, address._toCol) + ",";
                }
                if (address._toCol > addressToShift._toCol)
                {
                    newAddress += ExcelCellBase.GetAddress(address._fromRow, addressToShift._toCol + 1, address._toRow, address._toCol) + ",";
                }
                return new ExcelAddressBase(newAddress.Substring(0, newAddress.Length - 1));
            }
            else if (collide != ExcelAddressBase.eAddressCollition.No)
            {
                return ShiftAddress(address, range, shift);
            }
        }
        return address;
    }

    private static ExcelAddressBase ShiftAddress(ExcelAddressBase address, ExcelAddressBase range, eShiftTypeInsert shift)
    {
        if (shift == eShiftTypeInsert.Down)
        {
            return address.AddRow(range._fromRow, range.Rows);
        }
        else
        {
            return address.AddColumn(range._fromCol, range.Columns);
        }
    }

    private static void InsertPivottableAddress(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress)
    {
        foreach (ExcelPivotTable? ptbl in ws.PivotTables)
        {
            if (shift == eShiftTypeInsert.Down)
            {
                if (ptbl.Address._fromCol >= range._fromCol && ptbl.Address._toCol <= range._toCol)
                {
                    ptbl.Address = ptbl.Address.AddRow(range._fromRow, range.Rows);
                }
            }
            else
            {
                if (ptbl.Address._fromRow >= range._fromRow && ptbl.Address._toRow <= range._toRow)
                {
                    ptbl.Address = ptbl.Address.AddColumn(range._fromCol, range.Columns);
                }
            }

            if (ptbl.CacheDefinition.SourceRange.Worksheet == ws)
            {
                ExcelRangeBase? address = ptbl.CacheDefinition.SourceRange;
                if (shift == eShiftTypeInsert.Down)
                {
                    if (address._fromCol >= range._fromCol && address._toCol <= range._toCol)
                    {
                        ptbl.CacheDefinition.SourceRange = ws.Cells[address.AddRow(range._fromRow, range.Rows).Address];
                    }
                }
                else
                {
                    if (address._fromRow >= range._fromRow && address._toRow <= range._toRow)
                    {
                        ptbl.CacheDefinition.SourceRange = ws.Cells[address.AddColumn(range._fromCol, range.Columns).Address];
                    }
                }
            }
        }
    }

    private static void InsertTableAddress(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress)
    {
        foreach (ExcelTable? tbl in ws.Tables)
        {
            if (shift == eShiftTypeInsert.Down)
            {
                if (tbl.Address._fromCol >= range._fromCol && tbl.Address._toCol <= range._toCol)
                {
                    tbl.Address = tbl.Address.AddRow(range._fromRow, range.Rows);
                }
            }
            else
            {
                if (tbl.Address._fromRow >= range._fromRow && tbl.Address._toRow <= range._toRow)
                {
                    tbl.Address = tbl.Address.AddColumn(range._fromCol, range.Columns);
                }
            }

            //Update CalculatedColumnFormula
            ExcelAddressBase? address = tbl.Address.Intersect(range);
            foreach (ExcelTableColumn? col in tbl.Columns)
            {
                if (string.IsNullOrEmpty(col.CalculatedColumnFormula) == false)
                {
                    string? cf = ExcelCellBase.UpdateFormulaReferences(col.CalculatedColumnFormula, range, effectedAddress, shift, ws.Name, ws.Name);
                    col.SetFormula(cf);
                    if (address != null && tbl.Address._fromCol+col.Position-1 >= effectedAddress._fromCol)
                    {
                        int fromRow = tbl.ShowHeader && address._fromRow == tbl.Address._fromRow ? address._fromRow + 1 : address._fromRow;
                        int toRow = tbl.ShowTotal ? address._toRow - 1 : address._toRow;
                        int colNo = shift == eShiftTypeInsert.Right ? tbl.Address._fromCol + col.Position + range.Columns : tbl.Address._fromCol + col.Position;
                        col.SetFormulaCells(fromRow, toRow, colNo);
                    }
                }
            }
        }
    }

    private static List<int> GetStylesForRange(ExcelRangeBase range, eShiftTypeInsert shift)
    {
        List<int>? list = new List<int>();
        if (shift == eShiftTypeInsert.Down)
        {
            for (int i = 0; i < range.Columns; i++)
            {
                if (range._fromRow == 1)
                {
                    list.Add(0);
                }
                else
                {
                    list.Add(range.Offset(-1, i).StyleID);
                }
            }
        }
        else
        {
            for (int i = 0; i < range.Rows; i++)
            {
                if (range._fromCol == 1)
                {
                    list.Add(0);
                }
                else
                {
                    list.Add(range.Offset(i, -1).StyleID);
                }
            }
        }
        return list;
    }

    private static void SetStylesForRange(ExcelRangeBase range, eShiftTypeInsert shift, List<int> list)
    {
        if (shift == eShiftTypeInsert.Down)
        {
            for (int i = 0; i < range.Columns; i++)
            {
                range.Offset(0, i, range.Rows, 1).StyleID = list[i];
            }
        }
        else
        {
            for (int i = 0; i < range.Rows; i++)
            {

                range.Offset(i, 0, 1, range.Columns).StyleID = list[i];
            }
        }
    }

    private static ExcelAddressBase GetAffectedRange(ExcelRangeBase range, eShiftTypeInsert shift, int? start = null)
    {
        if (shift == eShiftTypeInsert.Down)
        {
            return new ExcelAddressBase(start ?? range._fromRow, range._fromCol, ExcelPackage.MaxRows, range._toCol);
        }
        else if (shift == eShiftTypeInsert.Right)
        {
            return new ExcelAddressBase(range._fromRow, start ?? range._fromCol, range._toRow, ExcelPackage.MaxColumns);
        }
        else if (shift == eShiftTypeInsert.EntireColumn)
        {
            return new ExcelAddressBase(1, range._fromCol, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        else
        {
            return new ExcelAddressBase(range._fromRow, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
    }

    private static void CopyStylesFromColumn(ExcelWorksheet ws, int columnFrom, int columns, int copyStylesFromColumn)
    {
        //Copy style from another column?
        if (copyStylesFromColumn > 0)
        {
            if (copyStylesFromColumn >= columnFrom)
            {
                copyStylesFromColumn += columns;
            }

            //Get styles to a cached list, 
            List<int[]>? l = new List<int[]>();
            CellStoreEnumerator<ExcelValue>? sce = new CellStoreEnumerator<ExcelValue>(ws._values, 0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
            lock (sce)
            {
                while (sce.Next())
                {
                    if (sce.Value._styleId == 0)
                    {
                        continue;
                    }

                    l.Add(new int[] { sce.Row, sce.Value._styleId });
                }
            }

            //Set the style id's from the list.
            foreach (int[]? sc in l)
            {
                for (int c = 0; c < columns; c++)
                {
                    if (sc[0] == 0)
                    {
                        ExcelColumn? col = ws.Column(columnFrom + c);   //Create the column
                        col.StyleID = sc[1];
                    }
                    else
                    {
                        ws.SetStyleInner(sc[0], columnFrom + c, sc[1]);
                    }
                }
            }
            int newOutlineLevel = ws.Column(copyStylesFromColumn).OutlineLevel;
            for (int c = 0; c < columns; c++)
            {
                ws.Column(columnFrom + c).OutlineLevel = newOutlineLevel;
            }
        }
    }

    private static void AdjustColumns(ExcelWorksheet ws, int columnFrom, int columns)
    {
        CellStoreEnumerator<ExcelValue>? csec = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
        List<ExcelColumn>? lst = new List<ExcelColumn>();
        foreach (ExcelValue val in csec)
        {
            object? col = val._value;
            if (col is ExcelColumn)
            {
                lst.Add((ExcelColumn)col);
            }
        }

        for (int i = lst.Count - 1; i >= 0; i--)
        {
            ExcelColumn? c = lst[i];
            if (c._columnMin >= columnFrom)
            {
                if (c._columnMin + columns <= ExcelPackage.MaxColumns)
                {
                    c._columnMin += columns;
                }
                else
                {
                    c._columnMin = ExcelPackage.MaxColumns;
                }

                if (c._columnMax + columns <= ExcelPackage.MaxColumns)
                {
                    c._columnMax += columns;
                }
                else
                {
                    c._columnMax = ExcelPackage.MaxColumns;
                }
            }
            else if (c._columnMax >= columnFrom)
            {
                int cc = c._columnMax - columnFrom;
                c._columnMax = columnFrom - 1;
                ws.CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
            }
        }
    }
    private static void AdjustFormulasInsert(ExcelRangeBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift)
    {
        //Adjust formulas
        foreach (ExcelWorksheet? ws in range._workbook.Worksheets)
        {
            string? workSheetName = range.Worksheet.Name;
            int rowFrom = range._fromRow;
            int columnFrom = range._fromCol;

            foreach (ExcelWorksheet.Formulas? f in ws._sharedFormulas.Values)
            {
                if (workSheetName == ws.Name)
                {
                    ExcelAddressBase? a = new ExcelAddressBase(f.Address);
                    ExcelAddressBase.eAddressCollition c = effectedAddress.Collide(a);
                    if (c == ExcelAddressBase.eAddressCollition.Partly && (effectedAddress._fromCol > a._fromCol || effectedAddress._toCol < a._toCol))
                    {
                        throw new Exception("Invalid shared formula"); //This should never happend!
                    }
                    if (f.StartCol >= columnFrom && c != ExcelAddressBase.eAddressCollition.No )
                    {
                        if(shift == eShiftTypeInsert.Down || shift == eShiftTypeInsert.EntireRow)
                        {
                            int rows = range.Rows;
                            if (f.StartRow >= rowFrom)
                            {
                                f.StartRow += rows;
                            }

                            if (a._fromRow >= rowFrom)
                            {
                                a._fromRow += rows;
                                a._toRow += rows;
                            }
                            else if (a._toRow >= rowFrom)
                            {
                                a._toRow += rows;
                            }
                        }
                        else
                        {
                            int cols = range.Columns;
                            if (f.StartCol >= columnFrom)
                            {
                                f.StartCol += cols;
                            }

                            if (a._fromCol >= columnFrom)
                            {
                                a._fromCol += cols;
                                a._toCol += cols;
                            }
                            else if (a._toCol >= columnFrom)
                            {
                                a._toCol += cols;
                            }
                        }
                        f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, range, effectedAddress, shift, ws.Name, workSheetName);
                    }
                }
                else if (f.Formula.Contains(workSheetName))
                {
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, range, effectedAddress, shift, ws.Name, workSheetName);
                }
                if (f.FormulaType == ExcelWorksheet.FormulaType.DataTable)
                {
                    if (string.IsNullOrEmpty(f.R1CellAddress) == false)
                    {
                        //var c1 = ExcelCellBase.Insert(f.Address, range);                            
                    }
                }
            }

            CellStoreEnumerator<object>? cse = new CellStoreEnumerator<object>(ws._formulas);
            while (cse.Next())
            {
                if (cse.Value is string v)
                {
                    if (workSheetName == ws.Name)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, range, effectedAddress, shift, ws.Name, workSheetName);
                    }
                    else if (v.Contains(workSheetName))
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, range, effectedAddress, shift, ws.Name, workSheetName);
                    }
                }
            }
        }
    }

    private static void FixFormulasInsertRow(ExcelWorksheet ws, int rowFrom, int rows, int columnFrom = 0, int columnTo = ExcelPackage.MaxColumns)
    {
        SourceCodeTokenizer? sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);

        //Adjust formulas
        foreach (ExcelWorksheet? wsToUpdate in ws.Workbook.Worksheets)
        {
            foreach (ExcelWorksheet.Formulas? f in wsToUpdate._sharedFormulas.Values)
            {
                if (ws.Name == wsToUpdate.Name)
                {
                    if (f.StartCol >= columnFrom)
                    {
                        if (f.StartRow >= rowFrom)
                        {
                            f.StartRow += rows;
                        }

                        ExcelAddressBase? a = new ExcelAddressBase(f.Address);
                        if (a._fromRow >= rowFrom)
                        {
                            a._fromRow += rows;
                            a._toRow += rows;
                        }
                        else if (a._toRow >= rowFrom)
                        {
                            a._toRow += rows;
                        }
                        f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name);
                    }
                }
                else if (f.Formula.Contains(ws.Name))
                {
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, columnFrom, wsToUpdate.Name, ws.Name);
                }
            }

            CellStoreEnumerator<object>? cse = new CellStoreEnumerator<object>(wsToUpdate._formulas);
            while (cse.Next())
            {
                if (cse.Value is string v)
                {
                    if (ws.Name == wsToUpdate.Name)
                    {
                        IEnumerable<Token>? tokens = GetTokens(wsToUpdate, cse.Row, cse.Column, v);
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name, false, false, tokens);
                    }
                    else if (v.Contains(ws.Name))
                    {
                        IEnumerable<Token>? tokens = GetTokens(wsToUpdate, cse.Row, cse.Column, v);
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name, false, false, tokens);
                    }
                    if(v!=cse.Value.ToString())
                    {
                        wsToUpdate._formulaTokens.SetValue(cse.Row, cse.Column, null);
                    }
                }
            }
        }
    }

    private static SourceCodeTokenizer _sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
    private static IEnumerable<Token> GetTokens(ExcelWorksheet ws, int row, int column, string formula)
    {
        return string.IsNullOrEmpty(formula) ? 
                   new List<Token>() : 
                   (List<Token>)_sct.Tokenize(formula, ws.Name);
    }
    private static void FixFormulasInsertColumn(ExcelWorksheet ws, int columnFrom, int columns)
    {
        foreach (ExcelWorksheet? wsToUpdate in ws.Workbook.Worksheets)
        {
            foreach (ExcelWorksheet.Formulas? f in wsToUpdate._sharedFormulas.Values)
            {
                if (ws.Name == wsToUpdate.Name)
                {
                    if (f.StartCol >= columnFrom)
                    {
                        f.StartCol += columns;
                    }

                    ExcelAddressBase? a = new ExcelAddressBase(f.Address);
                    if (a._fromCol >= columnFrom)
                    {
                        a._fromCol += columns;
                        a._toCol += columns;
                    }
                    else if (a._toCol >= columnFrom)
                    {
                        a._toCol += columns;
                    }

                    f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                }
                else if (f.Formula.Contains(ws.Name))
                {
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                }
            }

            CellStoreEnumerator<object>? cse = new CellStoreEnumerator<object>(wsToUpdate._formulas);
            while (cse.Next())
            {
                if (cse.Value is string v)
                {
                    if (ws.Name == wsToUpdate.Name)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                    }
                    else if (v.Contains(ws.Name))
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                    }
                }
            }
        }
    }
    private static void ValidateInsertColumn(ExcelWorksheet ws, int columnFrom, int columns, int rowFrom = 1, int rows = ExcelPackage.MaxRows)
    {
        ws.CheckSheetTypeAndNotDisposed();
        ExcelAddressBase? d = ws.Dimension;

        if (columnFrom < 1)
        {
            throw new ArgumentOutOfRangeException("columnFrom can't be lesser that 1");
        }
            
        //Check that cells aren't shifted outside the boundries.
        if (d != null && d.End.Column > columnFrom && d.End.Column + columns > ExcelPackage.MaxColumns)
        {
            throw new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet.");
        }

        ExcelAddressBase? insertRange = new ExcelAddressBase(rowFrom, columnFrom, rowFrom + rows - 1, columnFrom + columns - 1);
        FormulaDataTableValidation.HasPartlyFormulaDataTable(ws, insertRange, false, "Can't insert a part of a data table function");
    }
    #region private methods
    private static void ValidateInsertRow(ExcelWorksheet ws, int rowFrom, int rows, int columnFrom = 1, int columns = ExcelPackage.MaxColumns)
    {
        ws.CheckSheetTypeAndNotDisposed();
        ExcelAddressBase? d = ws.Dimension;

        if (rowFrom < 1)
        {
            throw new ArgumentOutOfRangeException("rowFrom can't be lesser that 1");
        }

        //Check that cells aren't shifted outside the boundries
        if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
        {
            throw new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet.");
        }

        ExcelAddressBase? insertRange = new ExcelAddressBase(rowFrom, columnFrom, rowFrom + rows - 1, columnFrom + columns - 1);
        FormulaDataTableValidation.HasPartlyFormulaDataTable(ws, insertRange, false, "Can't insert into a part of a data table function");
    }
    internal static void InsertCellStores(ExcelWorksheet ws, int rowFrom, int columnFrom, int rows, int columns, int columnTo = ExcelPackage.MaxColumns)
    {
        ws._values.Insert(rowFrom, columnFrom, rows, columns);
        ws._formulas.Insert(rowFrom, columnFrom, rows, columns);
        ws._formulaTokens.Insert(rowFrom, columnFrom, rows, columns);
        ws._commentsStore.Insert(rowFrom, columnFrom, rows, columns);
        ws._threadedCommentsStore.Insert(rowFrom, columnFrom, rows, columns);
        ws._hyperLinks.Insert(rowFrom, columnFrom, rows, columns);
        ws._dataValidationsStore.Insert(rowFrom, columnFrom, rows, columns);
        ws._flags.Insert(rowFrom, columnFrom, rows, columns);
        ws._metadataStore.Insert(rowFrom, columnFrom, rows, columns);
        ws._vmlDrawings?._drawingsCellStore.Insert(rowFrom, columnFrom, rows, columns);
        ws.MergedCells._cells.Insert(rowFrom, columnFrom, rows, columns);

        if (rows == 0 || columns == 0)
        {
            ws.Comments.Insert(rowFrom, columnFrom, rows, columns);
            ws.ThreadedComments.Insert(rowFrom, columnFrom, rows, columns);
            ws._names.Insert(rowFrom, columnFrom, rows, columns, 0, columnTo);
            ws.Workbook.Names.Insert(rowFrom, columnFrom, rows, columns, n => n.Worksheet == ws, 0, columnTo);
        }
        else
        {
            ws.Comments.Insert(rowFrom, columnFrom, rows, 0, 0, columnTo);
            ws.ThreadedComments.Insert(rowFrom, columnFrom, rows, 0, 0, columnTo);
            ws._names.Insert(rowFrom, columnFrom, rows, 0, columnFrom, columnTo);
            ws.Workbook.Names.Insert(rowFrom, columnFrom, rows, 0, n => n.Worksheet == ws, columnFrom, columnTo);
        }
    }
    internal static void InsertCellStoreShiftRight(ExcelWorksheet ws, ExcelAddressBase fromAddress)
    {
        ws._values.InsertShiftRight(fromAddress);
        ws._formulas.InsertShiftRight(fromAddress);
        ws._formulaTokens.InsertShiftRight(fromAddress);
        ws._commentsStore.InsertShiftRight(fromAddress);
        ws._threadedCommentsStore.InsertShiftRight(fromAddress);
        ws._hyperLinks.InsertShiftRight(fromAddress);
        ws._dataValidationsStore.InsertShiftRight(fromAddress);
        ws._flags.InsertShiftRight(fromAddress);
        ws._metadataStore.InsertShiftRight(fromAddress);
        ws._vmlDrawings?._drawingsCellStore.InsertShiftRight(fromAddress);
        ws.MergedCells._cells.InsertShiftRight(fromAddress); 

        ws.Comments.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, fromAddress._toRow, fromAddress._toCol);
        ws.ThreadedComments.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, fromAddress._toRow, fromAddress._toCol);
        ws._names.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, fromAddress._fromRow, fromAddress._toRow);
        ws.Workbook.Names.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, n => n.Worksheet == ws, fromAddress._fromRow, fromAddress._toRow);
    }

    private static void CopyFromStyleRow(ExcelWorksheet ws, int rowFrom, int rows, int copyStylesFromRow)
    {
        if (copyStylesFromRow >= rowFrom)
        {
            copyStylesFromRow += rows;
        }

        //Copy style from style row
        using (CellStoreEnumerator<ExcelValue>? cseS = new CellStoreEnumerator<ExcelValue>(ws._values, copyStylesFromRow, 0, copyStylesFromRow, ExcelPackage.MaxColumns))
        {
            while (cseS.Next())
            {
                if (cseS.Value._styleId == 0)
                {
                    continue;
                }

                for (int r = 0; r < rows; r++)
                {
                    ws.SetStyleInner(rowFrom + r, cseS.Column, cseS.Value._styleId);
                }
            }
        }

        //Copy outline
        int styleRowOutlineLevel = ws.Row(copyStylesFromRow).OutlineLevel;
        for (int r = rowFrom; r < rowFrom + rows; r++)
        {
            ws.Row(r).OutlineLevel = styleRowOutlineLevel;
        }
    }
    private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
    {
        XmlNode? node = tbl.Columns[0].TopNode.ParentNode;
        int ix = columnFrom - tbl.Address.Start.Column - 1;
        XmlNode? insPos = node.ChildNodes[ix];
        ix += 2;
        for (int i = 0; i < columns; i++)
        {
            string? name =
                tbl.Columns.GetUniqueName(string.Format("Column{0}",
                                                        (ix++).ToString(CultureInfo.InvariantCulture)));
            XmlElement tableColumn =
                (XmlElement)tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", ExcelPackage.schemaMain);
            tableColumn.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
            tableColumn.SetAttribute("name", name);
            insPos = node.InsertAfter(tableColumn, insPos);
        } //Create tbl Column
        tbl._cols = new ExcelTableColumnCollection(tbl);
    }

}
#endregion