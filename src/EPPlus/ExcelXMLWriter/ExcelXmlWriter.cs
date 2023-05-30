/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2023       EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml.Linq;
using static OfficeOpenXml.ExcelWorksheet;
using OfficeOpenXml.XMLWritingEncoder;

namespace OfficeOpenXml.ExcelXMLWriter;

internal class ExcelXmlWriter
{
    ExcelWorksheet _ws;
    ExcelPackage _package;
    private Dictionary<int, int> columnStyles;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="package"></param>
    public ExcelXmlWriter(ExcelWorksheet worksheet, ExcelPackage package)
    {
        this._ws = worksheet;
        this._package = package;
    }

    /// <summary>
    /// Replaces placeholder nodes by writing the system's held information
    /// </summary>
    /// <param name="sw">The streamwriter file info is written to</param>
    /// <param name="xml">The original XML</param>
    /// <param name="startOfNode">Start position of the current node</param>
    /// <param name="endOfNode">End position of the current node</param>
    internal void WriteNodes(StreamWriter sw, string xml, ref int startOfNode, ref int endOfNode)
    {
        string? prefix = this._ws.GetNameSpacePrefix();

        FindNodePositionAndClearItInit(sw, xml, "cols", ref startOfNode, ref endOfNode);
        this.UpdateColumnData(sw, prefix);

        FindNodePositionAndClearIt(sw, xml, "sheetData", ref startOfNode, ref endOfNode);
        this.UpdateRowCellData(sw, prefix);

        FindNodePositionAndClearIt(sw, xml, "mergeCells", ref startOfNode, ref endOfNode);
        this._ws._mergedCells.CleanupMergedCells();

        if (this._ws._mergedCells.Count > 0)
        {
            this.UpdateMergedCells(sw, prefix);
        }

        if (this._ws.GetNode("d:dataValidations") != null)
        {
            FindNodePositionAndClearIt(sw, xml, "dataValidations", ref startOfNode, ref endOfNode);

            if (this._ws.DataValidations.Count > 0)
            {
                sw.Write(this.UpdateDataValidation(prefix));
            }
        }

        FindNodePositionAndClearIt(sw, xml, "hyperlinks", ref startOfNode, ref endOfNode);
        this.UpdateHyperLinks(sw, prefix);

        FindNodePositionAndClearIt(sw, xml, "rowBreaks", ref startOfNode, ref endOfNode);
        this.UpdateRowBreaks(sw, prefix);

        FindNodePositionAndClearIt(sw, xml, "colBreaks", ref startOfNode, ref endOfNode);
        this.UpdateColBreaks(sw, prefix);

        //Careful. Ensure that we only do appropriate extLst things when there are objects to operate on.
        //Creating an empty DataValidations Node in ExtLst for example generates a corrupt excelfile that passes validation tool checks.
        if (this._ws.GetNode("d:extLst") != null && this._ws.DataValidations.GetExtLstCount() != 0)
        {
            ExtLstHelper extLst = new ExtLstHelper(xml);
            FindNodePositionAndClearIt(sw, xml, "extLst", ref startOfNode, ref endOfNode);

            extLst.InsertExt(ExtLstUris.DataValidationsUri, this.UpdateExtLstDataValidations(prefix), "");

            sw.Write(extLst.GetWholeExtLst());
        }

        sw.Write(xml.Substring(endOfNode, xml.Length - endOfNode));
    }

    internal static void FindNodePositionAndClearItInit(StreamWriter sw, string xml, string nodeName, ref int start, ref int end)
    {
        start = end;
        GetBlock.Pos(xml, nodeName, ref start, ref end);

        sw.Write(xml.Substring(0, start));
    }

    internal static void FindNodePositionAndClearIt(StreamWriter sw, string xml, string nodeName, ref int start, ref int end)
    {
        int oldEnd = end;
        GetBlock.Pos(xml, nodeName, ref start, ref end);

        sw.Write(xml.Substring(oldEnd, start - oldEnd));
    }

    /// <summary>
    /// Inserts the cols collection into the XML document
    /// </summary>
    private void UpdateColumnData(StreamWriter sw, string prefix)
    {
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(this._ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
        bool first = true;

        while (cse.Next())
        {
            ExcelColumn? col = cse.Value._value as ExcelColumn;

            if (col == null)
            {
                continue;
            }

            if (first)
            {
                sw.Write($"<{prefix}cols>");
                first = false;
            }

            ExcelStyleCollection<ExcelXfs> cellXfs = this._package.Workbook.Styles.CellXfs;

            sw.Write($"<{prefix}col min=\"{col.ColumnMin}\" max=\"{col.ColumnMax}\"");

            if (col.Hidden == true)
            {
                sw.Write(" hidden=\"1\"");
            }
            else if (col.BestFit)
            {
                sw.Write(" bestFit=\"1\"");
            }

            sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", col.Width));

            if (col.OutlineLevel > 0)
            {
                sw.Write($" outlineLevel=\"{col.OutlineLevel}\" ");

                if (col.Collapsed)
                {
                    sw.Write(" collapsed=\"1\"");
                }
            }

            if (col.Phonetic)
            {
                sw.Write(" phonetic=\"1\"");
            }

            int styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;

            if (styleID > 0)
            {
                sw.Write($" style=\"{styleID}\"");
            }

            sw.Write("/>");
        }

        if (!first)
        {
            sw.Write($"</{prefix}cols>");
        }
    }

    /// <summary>
    /// Check all Shared formulas that the first cell has not been deleted.
    /// If so create a standard formula of all cells in the formula .
    /// </summary>
    private void FixSharedFormulas()
    {
        List<int>? remove = new List<int>();

        foreach (Formulas? f in this._ws._sharedFormulas.Values)
        {
            ExcelAddressBase? addr = new ExcelAddressBase(f.Address);
            object? shIx = this._ws._formulas.GetValue(addr._fromRow, addr._fromCol);

            if (!(shIx is int) || (shIx is int && (int)shIx != f.Index))
            {
                for (int row = addr._fromRow; row <= addr._toRow; row++)
                {
                    for (int col = addr._fromCol; col <= addr._toCol; col++)
                    {
                        if (!(addr._fromRow == row && addr._fromCol == col))
                        {
                            object? fIx = this._ws._formulas.GetValue(row, col);

                            if (fIx is int && (int)fIx == f.Index)
                            {
                                this._ws._formulas.SetValue(row, col, f.GetFormula(row, col, this._ws.Name));
                            }
                        }
                    }
                }

                remove.Add(f.Index);
            }
        }

        remove.ForEach(i => this._ws._sharedFormulas.Remove(i));
    }

    // get StyleID without cell style for UpdateRowCellData
    internal int GetStyleIdDefaultWithMemo(int row, int col)
    {
        int v = 0;

        if (this._ws.ExistsStyleInner(row, 0, ref v)) //First Row
        {
            return v;
        }
        else // then column
        {
            if (!this.columnStyles.ContainsKey(col))
            {
                if (this._ws.ExistsStyleInner(0, col, ref v))
                {
                    this.columnStyles.Add(col, v);
                }
                else
                {
                    int r = 0,
                        c = col;

                    if (this._ws._values.PrevCell(ref r, ref c))
                    {
                        ExcelValue val = this._ws._values.GetValue(0, c);
                        ExcelColumn? column = (ExcelColumn)val._value;

                        if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                        {
                            this.columnStyles.Add(col, val._styleId);
                        }
                        else
                        {
                            this.columnStyles.Add(col, 0);
                        }
                    }
                    else
                    {
                        this.columnStyles.Add(col, 0);
                    }
                }
            }

            return this.columnStyles[col];
        }
    }

    private object GetFormulaValue(object v, string prefix)
    {
        if (v != null && v.ToString() != "")
        {
            return $"<{prefix}v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(v, this._ws.Workbook.Date1904))}</{prefix}v>";
        }
        else
        {
            return "";
        }
    }

    private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row, string prefix)
    {
        if (prevRow != -1)
        {
            _ = cache.Append($"</{prefix}row>");
        }

        //ulong rowID = ExcelRow.GetRowID(SheetID, row);
        _ = cache.Append($"<{prefix}row r=\"{row}\"");
        RowInternal currRow = this._ws.GetValueInner(row, 0) as RowInternal;

        if (currRow != null)
        {
            // if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
            if (currRow.Hidden == true)
            {
                _ = cache.Append(" hidden=\"1\"");
            }

            if (currRow.Height >= 0)
            {
                _ = cache.AppendFormat(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));

                if (currRow.CustomHeight)
                {
                    _ = cache.Append(" customHeight=\"1\"");
                }
            }

            if (currRow.OutlineLevel > 0)
            {
                _ = cache.AppendFormat(" outlineLevel =\"{0}\"", currRow.OutlineLevel);

                if (currRow.Collapsed)
                {
                    _ = cache.Append(" collapsed=\"1\"");
                }
            }

            if (currRow.Phonetic)
            {
                _ = cache.Append(" ph=\"1\"");
            }
        }

        int s = this._ws.GetStyleInner(row, 0);

        if (s > 0)
        {
            _ = cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID < 0 ? 0 : cellXfs[s].newID);
        }

        _ = cache.Append(">");
    }

    private static string GetDataTableAttributes(Formulas f)
    {
        string? attributes = " ";

        if (f.IsDataTableRow)
        {
            attributes += "dtr=\"1\" ";
        }

        if (f.DataTableIsTwoDimesional)
        {
            attributes += "dt2D=\"1\" ";
        }

        if (f.FirstCellDeleted)
        {
            attributes += "del1=\"1\" ";
        }

        if (f.SecondCellDeleted)
        {
            attributes += "del2=\"1\" ";
        }

        if (string.IsNullOrEmpty(f.R1CellAddress) == false)
        {
            attributes += $"r1=\"{f.R1CellAddress}\" ";
        }

        if (string.IsNullOrEmpty(f.R2CellAddress) == false)
        {
            attributes += $"r2=\"{f.R2CellAddress}\" ";
        }

        return attributes;
    }

    /// <summary>
    /// Insert row and cells into the XML document
    /// </summary>
    private void UpdateRowCellData(StreamWriter sw, string prefix)
    {
        ExcelStyleCollection<ExcelXfs> cellXfs = this._package.Workbook.Styles.CellXfs;

        int row = -1;
        string mdAttr = "";
        string mdAttrForFTag = "";
        string? sheetDataTag = prefix + "sheetData";
        string? cTag = prefix + "c";
        string? fTag = prefix + "f";
        string? vTag = prefix + "v";

        //_ = new StringBuilder();
        Dictionary<string, ExcelWorkbook.SharedStringItem>? ss = this._package.Workbook._sharedStrings;
        StringBuilder? cache = new StringBuilder();
        _ = cache.Append($"<{sheetDataTag}>");

        this.FixSharedFormulas(); //Fixes Issue #32

        bool hasMd = this._ws._metadataStore.HasValues;
        this.columnStyles = new Dictionary<int, int>();
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(this._ws._values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);

        while (cse.Next())
        {
            if (cse.Column > 0)
            {
                ExcelValue val = cse.Value;
                int styleID = cellXfs[val._styleId == 0 ? this.GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId].newID;
                styleID = styleID < 0 ? 0 : styleID;

                //Add the row element if it's a new row
                if (cse.Row != row)
                {
                    this.WriteRow(cache, cellXfs, row, cse.Row, prefix);
                    row = cse.Row;
                }

                object v = val._value;
                object formula = this._ws._formulas.GetValue(cse.Row, cse.Column);

                if (hasMd)
                {
                    mdAttr = "";

                    if (this._ws._metadataStore.Exists(cse.Row, cse.Column))
                    {
                        MetaDataReference md = this._ws._metadataStore.GetValue(cse.Row, cse.Column);

                        if (md.cm > 0)
                        {
                            mdAttr = $" cm=\"{md.cm}\"";
                        }

                        if (md.vm > 0)
                        {
                            mdAttr += $" vm=\"{md.vm}\"";
                        }
                    }
                }

                if (formula is int sfId)
                {
                    if (!this._ws._sharedFormulas.ContainsKey(sfId))
                    {
                        throw new
                            InvalidDataException($"SharedFormulaId {sfId} not found on Worksheet {this._ws.Name} cell {cse.CellAddress}, SharedFormulas Count {this._ws._sharedFormulas.Count}");
                    }

                    Formulas? f = this._ws._sharedFormulas[sfId];

                    //Set calc attributes for array formula. We preserve them from load only at this point.
                    if (hasMd)
                    {
                        mdAttrForFTag = "";

                        if (this._ws._metadataStore.Exists(cse.Row, cse.Column))
                        {
                            MetaDataReference md = this._ws._metadataStore.GetValue(cse.Row, cse.Column);

                            if (md.aca)
                            {
                                mdAttrForFTag = $" aca=\"1\"";
                            }

                            if (md.ca)
                            {
                                mdAttrForFTag += $" ca=\"1\"";
                            }
                        }
                    }

                    if (f.Address.IndexOf(':') > 0)
                    {
                        if (f.StartCol == cse.Column && f.StartRow == cse.Row)
                        {
                            if (f.FormulaType == FormulaType.Array)
                            {
                                _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"array\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                            else if (f.FormulaType == FormulaType.DataTable)
                            {
                                string? dataTableAttributes = GetDataTableAttributes(f);
                                _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"dataTable\"{dataTableAttributes} {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}></{cTag}>");
                            }
                            else
                            {
                                _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{f.Address}\" t=\"shared\" si=\"{sfId}\" {mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                            }
                        }
                        else if (f.FormulaType == FormulaType.Array)
                        {
                            string fElement;

                            if (string.IsNullOrEmpty(mdAttrForFTag) == false)
                            {
                                fElement = $"<{fTag} {mdAttrForFTag}/>";
                            }
                            else
                            {
                                fElement = $"";
                            }

                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>{fElement}{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                        else if (f.FormulaType == FormulaType.DataTable)
                        {
                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                        else
                        {
                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><f t=\"shared\" si=\"{sfId}\" {mdAttrForFTag}/>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                    }
                    else
                    {
                        // We can also have a single cell array formula
                        if (f.FormulaType == FormulaType.Array)
                        {
                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}><{fTag} ref=\"{string.Format("{0}:{1}", f.Address, f.Address)}\" t=\"array\"{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                        else
                        {
                            _ = cache.Append($"<{cTag} r=\"{f.Address}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                            _ = cache.Append($"<{fTag}{mdAttrForFTag}>{ConvertUtil.ExcelEscapeAndEncodeString(f.Formula)}</{fTag}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                    }
                }
                else if (formula != null && formula.ToString() != "")
                {
                    _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v, true)}{mdAttr}>");
                    _ = cache.Append($"<{fTag}>{ConvertUtil.ExcelEscapeAndEncodeString(formula.ToString())}</{fTag}>{this.GetFormulaValue(v, prefix)}</{cTag}>");
                }
                else
                {
                    if (v == null && styleID > 0)
                    {
                        _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{mdAttr}/>");
                    }
                    else if (v != null)
                    {
                        if (v is IEnumerable enumResult && !(v is string))
                        {
                            IEnumerator? e = enumResult.GetEnumerator();

                            if (e.MoveNext() && e.Current != null)
                            {
                                v = e.Current;
                            }
                            else
                            {
                                v = string.Empty;
                            }
                        }

                        if ((TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan) && !(v is char))
                        {
                            //string sv = GetValueForXml(v);
                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\"{ConvertUtil.GetCellType(v)}{mdAttr}>");
                            _ = cache.Append($"{this.GetFormulaValue(v, prefix)}</{cTag}>");
                        }
                        else
                        {
                            //If for example a struct 
                            string? s = Convert.ToString(v) ?? (v.ToString() ?? "");

                            int ix;

                            if (!ss.ContainsKey(s))
                            {
                                ix = ss.Count;

                                ss.Add(s,
                                       new ExcelWorkbook.SharedStringItem()
                                       {
                                           isRichText = this._ws._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.RichText), pos = ix
                                       });
                            }
                            else
                            {
                                ix = ss[s].pos;
                            }

                            _ = cache.Append($"<{cTag} r=\"{cse.CellAddress}\" s=\"{styleID}\" t=\"s\"{mdAttr}>");
                            _ = cache.Append($"<{vTag}>{ix}</{vTag}></{cTag}>");
                        }
                    }
                }
            }
            else //ExcelRow
            {
                this.WriteRow(cache, cellXfs, row, cse.Row, prefix);
                row = cse.Row;
            }

            if (cache.Length > 0x600000)
            {
                sw.Write(cache.ToString());
                sw.Flush();
                cache.Length = 0;
            }
        }

        this.columnStyles = null;

        if (row != -1)
        {
            _ = cache.Append($"</{prefix}row>");
        }

        _ = cache.Append($"</{prefix}sheetData>");
        sw.Write(cache.ToString());
        sw.Flush();
    }

    /// <summary>
    /// Update merged cells
    /// </summary>
    /// <param name="sw">The writer</param>
    /// <param name="prefix">Namespace prefix for the main schema</param>
    private void UpdateMergedCells(StreamWriter sw, string prefix)
    {
        sw.Write($"<{prefix}mergeCells>");

        foreach (string address in this._ws._mergedCells.Distinct())
        {
            sw.Write($"<{prefix}mergeCell ref=\"{address}\" />");
        }

        sw.Write($"</{prefix}mergeCells>");
    }

    private void WriteDataValidationAttributes(StringBuilder cache, int i)
    {
        if (this._ws.DataValidations[i].ValidationType != null && this._ws.DataValidations[i].ValidationType.Type != eDataValidationType.Any)
        {
            _ = cache.Append($"type=\"{this._ws.DataValidations[i].ValidationType.TypeToXmlString()}\" ");
        }

        if (this._ws.DataValidations[i].ErrorStyle != ExcelDataValidationWarningStyle.undefined)
        {
            _ = cache.Append($"errorStyle=\"{this._ws.DataValidations[i].ErrorStyle.ToEnumString()}\" ");
        }

        if (this._ws.DataValidations[i].ImeMode != ExcelDataValidationImeMode.NoControl)
        {
            _ = cache.Append($"imeMode=\"{this._ws.DataValidations[i].ImeMode.ToEnumString()}\" ");
        }

        if (this._ws.DataValidations[i].Operator != 0)
        {
            _ = cache.Append($"operator=\"{this._ws.DataValidations[i].Operator.ToEnumString()}\" ");
        }

        //Note that if false excel does not write these properties out so we don't either.
        if (this._ws.DataValidations[i].AllowBlank == true)
        {
            _ = cache.Append($"allowBlank=\"1\" ");
        }

        if (this._ws.DataValidations[i] is ExcelDataValidationList)
        {
            if ((this._ws.DataValidations[i] as ExcelDataValidationList).HideDropDown == true)
            {
                _ = cache.Append($"showDropDown=\"1\" ");
            }
        }

        if (this._ws.DataValidations[i].ShowInputMessage == true)
        {
            _ = cache.Append($"showInputMessage=\"1\" ");
        }

        if (this._ws.DataValidations[i].ShowErrorMessage == true)
        {
            _ = cache.Append($"showErrorMessage=\"1\" ");
        }

        if (string.IsNullOrEmpty(this._ws.DataValidations[i].ErrorTitle) == false)
        {
            _ = cache.Append($"errorTitle=\"{this._ws.DataValidations[i].ErrorTitle.EncodeXMLAttribute()}\" ");
        }

        if (string.IsNullOrEmpty(this._ws.DataValidations[i].Error) == false)
        {
            _ = cache.Append($"error=\"{this._ws.DataValidations[i].Error.EncodeXMLAttribute()}\" ");
        }

        if (string.IsNullOrEmpty(this._ws.DataValidations[i].PromptTitle) == false)
        {
            _ = cache.Append($"promptTitle=\"{this._ws.DataValidations[i].PromptTitle.EncodeXMLAttribute()}\" ");
        }

        if (string.IsNullOrEmpty(this._ws.DataValidations[i].Prompt) == false)
        {
            _ = cache.Append($"prompt=\"{this._ws.DataValidations[i].Prompt.EncodeXMLAttribute()}\" ");
        }

        if (this._ws.DataValidations[i].InternalValidationType == InternalValidationType.DataValidation)
        {
            _ = cache.Append($"sqref=\"{this._ws.DataValidations[i].Address.ToString().Replace(",", " ")}\" ");
        }

        _ = cache.Append($"xr:uid=\"{this._ws.DataValidations[i].Uid}\"");

        _ = cache.Append(">");
    }

    private void WriteDataValidation(StringBuilder cache, string prefix, int i, string extNode = "")
    {
        _ = cache.Append($"<{prefix}dataValidation ");
        this.WriteDataValidationAttributes(cache, i);

        if (this._ws.DataValidations[i].ValidationType.Type != eDataValidationType.Any)
        {
            string endExtNode = "";

            if (extNode != "")
            {
                endExtNode = $"</{extNode}>";
                extNode = $"<{extNode}>";
            }

            switch (this._ws.DataValidations[i].ValidationType.Type)
            {
                case eDataValidationType.TextLength:
                case eDataValidationType.Whole:
                    ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>? intType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>;

                    WriteDataValidationFormulas(intType.Formula, intType.Formula2, cache, prefix, extNode, endExtNode, this._ws.DataValidations[i].Operator);

                    break;

                case eDataValidationType.Decimal:
                    ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>? decimalType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>;

                    WriteDataValidationFormulas(decimalType.Formula,
                                                decimalType.Formula2,
                                                cache,
                                                prefix,
                                                extNode,
                                                endExtNode,
                                                this._ws.DataValidations[i].Operator);

                    break;

                case eDataValidationType.List:
                    ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>? listType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>;

                    WriteDataValidationFormulaSingle(listType.Formula, cache, prefix, extNode, endExtNode);

                    break;

                case eDataValidationType.Time:
                    ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>? timeType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>;

                    WriteDataValidationFormulas(timeType.Formula, timeType.Formula2, cache, prefix, extNode, endExtNode, this._ws.DataValidations[i].Operator);

                    break;

                case eDataValidationType.DateTime:
                    ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>? dateTimeType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>;

                    WriteDataValidationFormulas(dateTimeType.Formula,
                                                dateTimeType.Formula2,
                                                cache,
                                                prefix,
                                                extNode,
                                                endExtNode,
                                                this._ws.DataValidations[i].Operator);

                    break;

                case eDataValidationType.Custom:
                    ExcelDataValidationWithFormula<IExcelDataValidationFormula>? customType =
                        this._ws.DataValidations[i] as ExcelDataValidationWithFormula<IExcelDataValidationFormula>;

                    WriteDataValidationFormulaSingle(customType.Formula, cache, prefix, extNode, endExtNode);

                    break;

                default:
                    throw new Exception("UNKNOWN TYPE IN WriteDataValidation");
            }

            if (extNode != "")
            {
                //write adress if extLst
                _ = cache.Append($"<xm:sqref>{this._ws.DataValidations[i].Address.ToString().Replace(",", " ")}</xm:sqref>");
            }
        }

        _ = cache.Append($"</{prefix}dataValidation>");
    }

    static void WriteDataValidationFormulaSingle(IExcelDataValidationFormula formula, StringBuilder cache, string prefix, string extNode, string endExtNode)
    {
        string string1 = ((ExcelDataValidationFormula)formula).GetXmlValue();
        string1 = ConvertUtil.ExcelEscapeAndEncodeString(string1);

        _ = cache.Append($"<{prefix}formula1>{extNode}{string1}{endExtNode}</{prefix}formula1>");
    }

    static void WriteDataValidationFormulas(IExcelDataValidationFormula formula1,
                                            IExcelDataValidationFormula formula2,
                                            StringBuilder cache,
                                            string prefix,
                                            string extNode,
                                            string endExtNode,
                                            ExcelDataValidationOperator dvOperator)
    {
        string string1 = ((ExcelDataValidationFormula)formula1).GetXmlValue();
        string string2 = ((ExcelDataValidationFormula)formula2).GetXmlValue();

        //Note that formula1 must be written even when string1 is empty
        string1 = ConvertUtil.ExcelEscapeAndEncodeString(string1);
        _ = cache.Append($"<{prefix}formula1>{extNode}{string1}{endExtNode}</{prefix}formula1>");

        if (!string.IsNullOrEmpty(string2) && (dvOperator == ExcelDataValidationOperator.between || dvOperator == ExcelDataValidationOperator.notBetween))
        {
            string2 = ConvertUtil.ExcelEscapeAndEncodeString(string2);
            _ = cache.Append($"<{prefix}formula2>{extNode}{string2}{endExtNode}</{prefix}formula2>");
        }
    }

    private StringBuilder UpdateDataValidation(string prefix, string extraAttribute = "")
    {
        StringBuilder? cache = new StringBuilder();
        InternalValidationType type;
        string extNode = "";

        if (extraAttribute == "")
        {
            _ = cache.Append($"<{prefix}dataValidations count=\"{this._ws.DataValidations.GetNonExtLstCount()}\">");
            type = InternalValidationType.DataValidation;
        }
        else
        {
            _ = cache.Append($"<{prefix}dataValidations {extraAttribute} count=\"{this._ws.DataValidations.GetExtLstCount()}\">");
            type = InternalValidationType.ExtLst;
            extNode = "xm:f";
        }

        for (int i = 0; i < this._ws.DataValidations.Count; i++)
        {
            if (this._ws.DataValidations[i].InternalValidationType == type)
            {
                this.WriteDataValidation(cache, prefix, i, extNode);
            }
        }

        _ = cache.Append($"</{prefix}dataValidations>");

        return cache;
    }

    /// <summary>
    /// Update xml with hyperlinks 
    /// </summary>
    /// <param name="sw">The stream</param>
    /// <param name="prefix">The namespace prefix for the main schema</param>
    private void UpdateHyperLinks(StreamWriter sw, string prefix)
    {
        Dictionary<string, string> hyps = new Dictionary<string, string>();
        CellStoreEnumerator<Uri>? cse = new CellStoreEnumerator<Uri>(this._ws._hyperLinks);
        bool first = true;

        while (cse.Next())
        {
            Uri? uri = this._ws._hyperLinks.GetValue(cse.Row, cse.Column);

            if (first && uri != null)
            {
                sw.Write($"<{prefix}hyperlinks>");
                first = false;
            }

            ExcelHyperLink? hl = uri as ExcelHyperLink;

            if (hl != null && !string.IsNullOrEmpty(hl.ReferenceAddress))
            {
                string? address = this._ws.Cells[cse.Row, cse.Column, cse.Row + hl.RowSpann, cse.Column + hl.ColSpann].Address;
                string? location = ExcelCellBase.GetFullAddress(SecurityElement.Escape(this._ws.Name), SecurityElement.Escape(hl.ReferenceAddress));
                string? display = string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"";
                string? tooltip = string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"";
                sw.Write($"<{prefix}hyperlink ref=\"{address}\" location=\"{location}\"{display}{tooltip}/>");
            }
            else if (uri != null)
            {
                Uri hyp;
                string target = "";
                ;

                if (hl != null)
                {
                    if (hl.Target != null && hl.OriginalString.StartsWith("Invalid:Uri", StringComparison.OrdinalIgnoreCase))
                    {
                        target = hl.Target;
                    }

                    hyp = hl.OriginalUri;
                }
                else
                {
                    hyp = uri;
                }

                if (hyps.ContainsKey(hyp.OriginalString) && string.IsNullOrEmpty(target))
                {
                }
                else
                {
                    ZipPackageRelationship relationship;

                    if (string.IsNullOrEmpty(target))
                    {
                        relationship = this._ws.Part.CreateRelationship(hyp, TargetMode.External, ExcelPackage.schemaHyperlink);
                    }
                    else
                    {
                        relationship = this._ws.Part.CreateRelationship(target, TargetMode.External, ExcelPackage.schemaHyperlink);
                    }

                    if (hl != null)
                    {
                        string? display = string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"";
                        string? toolTip = string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"";
                        sw.Write($"<{prefix}hyperlink ref=\"{ExcelCellBase.GetAddress(cse.Row, cse.Column)}\"{display}{toolTip} r:id=\"{relationship.Id}\"/>");
                    }
                    else
                    {
                        sw.Write($"<{prefix}hyperlink ref=\"{ExcelCellBase.GetAddress(cse.Row, cse.Column)}\" r:id=\"{relationship.Id}\"/>");
                    }
                }
            }
        }

        if (!first)
        {
            sw.Write($"</{prefix}hyperlinks>");
        }
    }

    private void UpdateRowBreaks(StreamWriter sw, string prefix)
    {
        StringBuilder breaks = new StringBuilder();
        int count = 0;
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(this._ws._values, 0, 0, ExcelPackage.MaxRows, 0);

        while (cse.Next())
        {
            RowInternal? row = cse.Value._value as RowInternal;

            if (row != null && row.PageBreak)
            {
                _ = breaks.AppendFormat($"<{prefix}brk id=\"{cse.Row}\" max=\"1048575\" man=\"1\"/>");
                count++;
            }
        }

        if (count > 0)
        {
            sw.Write(string.Format($"<{prefix}rowBreaks count=\"{count}\" manualBreakCount=\"{count}\">{breaks.ToString()}</rowBreaks>"));
        }
    }

    private void UpdateColBreaks(StreamWriter sw, string prefix)
    {
        StringBuilder breaks = new StringBuilder();
        int count = 0;
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(this._ws._values, 0, 0, 0, ExcelPackage.MaxColumns);

        while (cse.Next())
        {
            ExcelColumn? col = cse.Value._value as ExcelColumn;

            if (col != null && col.PageBreak)
            {
                _ = breaks.Append($"<{prefix}brk id=\"{cse.Column}\" max=\"16383\" man=\"1\"/>");
                count++;
            }
        }

        if (count > 0)
        {
            sw.Write($"<colBreaks count=\"{count}\" manualBreakCount=\"{count}\">{breaks.ToString()}</colBreaks>");
        }
    }

    /// <summary>
    /// ExtLst updater for DataValidations
    /// </summary>
    /// <param name="prefix"></param>
    /// <returns></returns>
    private string UpdateExtLstDataValidations(string prefix)
    {
        StringBuilder? cache = new StringBuilder();

        _ = cache.Append($"<ext xmlns:x14=\"{ExcelPackage.schemaMainX14}\" uri=\"{ExtLstUris.DataValidationsUri}\">");

        prefix = "x14:";
        _ = cache.Append(this.UpdateDataValidation(prefix, $"xmlns:xm=\"{ExcelPackage.schemaMainXm}\""));
        _ = cache.Append("</ext>");

        return cache.ToString();
    }
}