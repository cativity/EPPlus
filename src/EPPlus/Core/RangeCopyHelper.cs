/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation.Contracts;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core;

internal class RangeCopyHelper
{
    private class CopiedCell
    {
        internal int Row { get; set; }

        internal int Column { get; set; }

        internal object Value { get; set; }

        internal string Type { get; set; }

        internal object Formula { get; set; }

        internal int? StyleID { get; set; }

        internal Uri HyperLink { get; set; }

        internal ExcelComment Comment { get; set; }

        internal ExcelThreadedCommentThread ThreadedComment { get; set; }

        internal byte Flag { get; set; }

        internal ExcelWorksheet.MetaDataReference MetaData { get; set; }
    }

    private readonly ExcelRangeBase _sourceRange;
    private readonly ExcelRangeBase _destination;
    private readonly ExcelRangeCopyOptionFlags _copyOptions;
    Dictionary<ulong, CopiedCell> _copiedCells = new Dictionary<ulong, CopiedCell>();

    internal RangeCopyHelper(ExcelRangeBase sourceRange, ExcelRangeBase destination, ExcelRangeCopyOptionFlags copyOptions)
    {
        this._sourceRange = sourceRange;
        this._destination = destination;
        this._copyOptions = copyOptions;
    }

    internal void Copy()
    {
        this.GetCopiedValues();

        Dictionary<int, ExcelAddress> copiedMergedCells;

        if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeMergedCells))
        {
            copiedMergedCells = this.GetCopiedMergedCells();
        }
        else
        {
            copiedMergedCells = null;
        }

        this.ClearDestination();

        this.CopyValuesToDestination();

        if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeDataValidations))
        {
            this.CopyDataValidations();
        }

        if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting))
        {
            this.CopyConditionalFormatting();
        }

        if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeMergedCells))
        {
            this.CopyMergedCells(copiedMergedCells);
        }

        this.CopyFullColumn();
        this.CopyFullRow();
    }

    private void CopyDataValidations()
    {
        foreach (IExcelDataValidation? idv in this._sourceRange._worksheet.DataValidations)
        {
            if (idv is ExcelDataValidation dv)
            {
                string newAddress = "";

                if (dv.Address.Addresses == null)
                {
                    newAddress = this.HandelAddress(dv.Address);
                }
                else
                {
                    foreach (ExcelAddressBase? a in dv.Address.Addresses)
                    {
                        string? na = this.HandelAddress(a);

                        if (!string.IsNullOrEmpty(na))
                        {
                            if (string.IsNullOrEmpty(newAddress))
                            {
                                newAddress += na;
                            }
                            else
                            {
                                newAddress += "," + na;
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(newAddress) == false)
                {
                    if (this._sourceRange._worksheet == this._destination._worksheet)
                    {
                        dv.SetAddress(dv.Address + "," + newAddress);
                    }
                    else
                    {
                        this._destination._worksheet.DataValidations.AddCopyOfDataValidation(dv,
                                                                                             this._destination._worksheet,
                                                                                             new ExcelAddressBase(newAddress).AddressSpaceSeparated);
                    }
                }
            }
        }
    }

    private void CopyConditionalFormatting()
    {
        foreach (IExcelConditionalFormattingRule? cf in this._sourceRange._worksheet.ConditionalFormatting)
        {
            string newAddress = "";

            if (cf.Address.Addresses == null)
            {
                newAddress = this.HandelAddress(cf.Address);
            }
            else
            {
                foreach (ExcelAddressBase? a in cf.Address.Addresses)
                {
                    string? na = this.HandelAddress(a);

                    if (!string.IsNullOrEmpty(na))
                    {
                        if (string.IsNullOrEmpty(newAddress))
                        {
                            newAddress += na;
                        }
                        else
                        {
                            newAddress += "," + na;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(newAddress) == false)
            {
                if (this._sourceRange._worksheet == this._destination._worksheet)
                {
                    cf.Address = new ExcelAddress(cf.Address + "," + newAddress);
                }
                else
                {
                    this._destination._worksheet.ConditionalFormatting.AddFromXml(new ExcelAddress(newAddress), cf.PivotTable, cf.Node.OuterXml);

                    if (cf.Style.HasValue)
                    {
                        ExcelConditionalFormattingRule? destRule =
                            (ExcelConditionalFormattingRule)this._destination._worksheet.ConditionalFormatting[this._destination._worksheet
                                    .ConditionalFormatting.Count
                                - 1];

                        destRule.SetStyle((ExcelDxfStyleConditionalFormatting)cf.Style.Clone());
                    }
                }
            }
        }
    }

    private string HandelAddress(ExcelAddressBase cfAddress)
    {
        if (cfAddress.Collide(this._sourceRange) != eAddressCollition.No)
        {
            ExcelAddressBase? address = this._sourceRange.Intersect(cfAddress);
            int rowOffset = address._fromRow - this._sourceRange._fromRow;
            int colOffset = address._fromCol - this._sourceRange._fromCol;
            int fr = Math.Min(Math.Max(this._destination._fromRow + rowOffset, 1), ExcelPackage.MaxRows);
            int fc = Math.Min(Math.Max(this._destination._fromCol + colOffset, 1), ExcelPackage.MaxColumns);

            address = new ExcelAddressBase(fr,
                                           fc,
                                           Math.Min(fr + address.Rows - 1, ExcelPackage.MaxRows),
                                           Math.Min(fc + address.Columns - 1, ExcelPackage.MaxColumns));

            return address.Address;
        }

        return "";
    }

    private void GetCopiedValues()
    {
        ExcelWorksheet? worksheet = this._sourceRange._worksheet;

        _ = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues);
        bool includeStyles = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles);
        bool includeComments = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments);
        bool includeThreadedComments = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments);

        Dictionary<int, int> styleCashe = new Dictionary<int, int>();
        bool sameWorkbook = this._destination._worksheet.Workbook == this._sourceRange._worksheet.Workbook;

        this.AddValuesFormulasAndStyles(worksheet, includeStyles, styleCashe, sameWorkbook);

        if (includeComments)
        {
            this.AddComments(worksheet);
        }

        if (includeThreadedComments)
        {
            this.AddThreadedComments(worksheet);
        }
    }

    private void AddValuesFormulasAndStyles(ExcelWorksheet worksheet, bool includeStyles, Dictionary<int, int> styleCashe, bool sameWorkbook)
    {
        int styleId = 0;
        object o = null;
        byte flag = 0;
        Uri hl = null;

        bool includeValues = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues);
        bool includeFormulas = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas);
        bool includeHyperlinks = EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks);

        if (includeValues == false && includeHyperlinks == false && includeFormulas == false)
        {
            return;
        }

        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(worksheet._values,
                                                                                   this._sourceRange._fromRow,
                                                                                   this._sourceRange._fromCol,
                                                                                   this._sourceRange._toRow,
                                                                                   this._sourceRange._toCol);

        while (cse.Next())
        {
            int row = cse.Row;
            int col = cse.Column; //Issue 15070

            CopiedCell? cell = new CopiedCell
            {
                Row = this._destination._fromRow + (row - this._sourceRange._fromRow),
                Column = this._destination._fromCol + (col - this._sourceRange._fromCol),
            };

            if (includeValues)
            {
                cell.Value = cse.Value._value;
            }

            if (includeFormulas && worksheet._formulas.Exists(row, col, ref o))
            {
                if (o is int)
                {
                    cell.Formula = worksheet.GetFormula(cse.Row, cse.Column);

                    if (worksheet._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.ArrayFormula))
                    {
                        this._destination._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
                    }

                    // We currently don't copy CellFlags.DataTableFormula's, as Excel does not.
                }
                else
                {
                    cell.Formula = o;
                }
            }

            if (includeStyles && worksheet.ExistsStyleInner(row, col, ref styleId))
            {
                if (sameWorkbook)
                {
                    cell.StyleID = styleId;
                }
                else
                {
                    if (styleCashe.ContainsKey(styleId))
                    {
                        styleId = styleCashe[styleId];
                    }
                    else
                    {
                        int oldStyleID = styleId;
                        styleId = this._destination._worksheet.Workbook.Styles.CloneStyle(this._sourceRange._worksheet.Workbook.Styles, styleId);
                        styleCashe.Add(oldStyleID, styleId);
                    }

                    cell.StyleID = styleId;
                }
            }

            ExcelWorksheet.MetaDataReference md = new ExcelWorksheet.MetaDataReference();

            if (includeFormulas && worksheet._metadataStore.Exists(row, col, ref md))
            {
                cell.MetaData = md;
            }

            if (includeHyperlinks && worksheet._hyperLinks.Exists(row, col, ref hl))
            {
                cell.HyperLink = hl;
            }

            if (worksheet._flags.Exists(row, col, ref flag))
            {
                cell.Flag = flag;
            }

            this._copiedCells.Add(ExcelCellBase.GetCellId(0, row, col), cell);
        }
    }

    private void AddComments(ExcelWorksheet worksheet)
    {
        CellStoreEnumerator<int>? cse = new CellStoreEnumerator<int>(worksheet._commentsStore,
                                                                     this._sourceRange._fromRow,
                                                                     this._sourceRange._fromCol,
                                                                     this._sourceRange._toRow,
                                                                     this._sourceRange._toCol);

        while (cse.Next())
        {
            int row = cse.Row;
            int col = cse.Column; //Issue 15070
            ulong cellId = ExcelCellBase.GetCellId(0, row, col);
            CopiedCell cell;

            if (this._copiedCells.ContainsKey(cellId))
            {
                cell = this._copiedCells[cellId];
            }
            else
            {
                cell = new CopiedCell
                {
                    Row = this._destination._fromRow + (row - this._sourceRange._fromRow),
                    Column = this._destination._fromCol + (col - this._sourceRange._fromCol),
                };

                this._copiedCells.Add(cellId, cell);
            }

            cell.Comment = worksheet._comments[cse.Value];
        }
    }

    private void AddThreadedComments(ExcelWorksheet worksheet)
    {
        CellStoreEnumerator<int>? cse = new CellStoreEnumerator<int>(worksheet._threadedCommentsStore,
                                                                     this._sourceRange._fromRow,
                                                                     this._sourceRange._fromCol,
                                                                     this._sourceRange._toRow,
                                                                     this._sourceRange._toCol);

        while (cse.Next())
        {
            int row = cse.Row;
            int col = cse.Column; //Issue 15070
            ulong cellId = ExcelCellBase.GetCellId(0, row, col);
            CopiedCell cell;

            if (this._copiedCells.ContainsKey(cellId))
            {
                cell = this._copiedCells[cellId];
            }
            else
            {
                cell = new CopiedCell
                {
                    Row = this._destination._fromRow + (row - this._sourceRange._fromRow),
                    Column = this._destination._fromCol + (col - this._sourceRange._fromCol),
                };

                this._copiedCells.Add(cellId, cell);
            }

            cell.ThreadedComment = worksheet._threadedComments[cse.Value];
        }
    }

    private void CopyValuesToDestination()
    {
        int fromRow = this._sourceRange._fromRow;
        int fromCol = this._sourceRange._fromCol;

        foreach (CopiedCell? cell in this._copiedCells.Values)
        {
            if (EnumUtil.HasFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues)
                && EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
            {
                this._destination._worksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID ?? 0);
            }
            else if (EnumUtil.HasFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
            {
                this._destination._worksheet.SetValueInner(cell.Row, cell.Column, cell.Value);
            }
            else
            {
                this._destination._worksheet.SetValueStyleIdInner(cell.Row, cell.Column, cell.Value, cell.StyleID ?? 0);
            }

            if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas)
                && EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues)
                && cell.Formula != null)
            {
                cell.Formula = ExcelCellBase.UpdateFormulaReferences(cell.Formula.ToString(),
                                                                     this._destination._fromRow - fromRow,
                                                                     this._destination._fromCol - fromCol,
                                                                     0,
                                                                     0,
                                                                     this._destination.WorkSheetName,
                                                                     this._destination.WorkSheetName,
                                                                     true,
                                                                     true);

                this._destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
            }

            if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeHyperLinks) && cell.HyperLink != null)
            {
                this._destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
            }

            if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeThreadedComments) && cell.ThreadedComment != null)
            {
                bool differentPackages = this._destination._workbook != this._sourceRange._workbook;
                ExcelThreadedCommentThread? tc = this._destination.Worksheet.Cells[cell.Row, cell.Column].AddThreadedComment();

                foreach (ExcelThreadedComment? c in cell.ThreadedComment.Comments)
                {
                    if (differentPackages && this._destination._workbook.ThreadedCommentPersons[c.PersonId] == null)
                    {
                        ExcelThreadedCommentPerson? p = this._sourceRange._workbook.ThreadedCommentPersons[c.PersonId];
                        _ = this._destination._workbook.ThreadedCommentPersons.Add(p.DisplayName, p.UserId, p.ProviderId, p.Id);
                    }

                    tc.AddCommentFromXml((XmlElement)c.TopNode);
                }
            }
            else if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeComments) && cell.Comment != null)
            {
                CopyComment(this._destination, cell);
            }

            if (cell.Flag != 0)
            {
                this._destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
            }

            if ((EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeFormulas | ExcelRangeCopyOptionFlags.ExcludeValues)
                 && cell.MetaData.cm > 0)
                || cell.MetaData.vm > 0)
            {
                this._destination._worksheet._metadataStore.SetValue(cell.Row, cell.Column, cell.MetaData);
            }
        }
    }

    private static void CopyComment(ExcelRangeBase destination, CopiedCell cell)
    {
        ExcelComment? c = destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment.Text, cell.Comment.Author);
        int offsetCol = c.Column - cell.Comment.Column;
        int offsetRow = c.Row - cell.Comment.Row;
        XmlHelper.CopyElement((XmlElement)cell.Comment.TopNode, (XmlElement)c.TopNode, new string[] { "id", "spid" });

        if (c.From.Column + offsetCol >= 0)
        {
            c.From.Column += offsetCol;
            c.To.Column += offsetCol;
        }

        if (c.From.Row + offsetRow >= 0)
        {
            c.From.Row += offsetRow;
            c.To.Row += offsetRow;
        }

        c.Row = cell.Row - 1;
        c.Column = cell.Column - 1;

        c._commentHelper.TopNode.InnerXml = cell.Comment._commentHelper.TopNode.InnerXml;
        c.RichText = new Style.ExcelRichTextCollection(c._commentHelper.NameSpaceManager, c._commentHelper.GetNode("d:text"), destination._worksheet);

        //Add relation to image used for filling the comment
        if (cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Frame
            || cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Tile
            || cell.Comment.Fill.Style == Drawing.Vml.eVmlFillType.Pattern)
        {
            ExcelImage? img = cell.Comment.Fill.PatternPictureSettings.Image;

            if (img.ImageBytes != null)
            {
                _ = c.Fill.PatternPictureSettings.Image.SetImage(img.ImageBytes, img.Type ?? ePictureType.Jpg);
            }
        }
    }

    private void ClearDestination()
    {
        //Clear all existing cells; 
        int rows = this._sourceRange._toRow - this._sourceRange._fromRow + 1,
            cols = this._sourceRange._toCol - this._sourceRange._fromCol + 1;

        this._destination._worksheet.MergedCells.Clear(new ExcelAddressBase(this._destination._fromRow,
                                                                            this._destination._fromCol,
                                                                            this._destination._fromRow + rows - 1,
                                                                            this._destination._fromCol + cols - 1));

        if (EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeValues)
            && EnumUtil.HasNotFlag(this._copyOptions, ExcelRangeCopyOptionFlags.ExcludeStyles))
        {
            this._destination._worksheet._values.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        }

        this._destination._worksheet._formulas.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        this._destination._worksheet._metadataStore.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        this._destination._worksheet._hyperLinks.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        this._destination._worksheet._flags.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        this._destination._worksheet._commentsStore.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
        this._destination._worksheet._threadedCommentsStore.Clear(this._destination._fromRow, this._destination._fromCol, rows, cols);
    }

    private Dictionary<int, ExcelAddress> GetCopiedMergedCells()
    {
        ExcelWorksheet? worksheet = this._sourceRange._worksheet;
        Dictionary<int, ExcelAddress>? copiedMergedCells = new Dictionary<int, ExcelAddress>();

        //Merged cells
        CellStoreEnumerator<int>? csem = new CellStoreEnumerator<int>(worksheet.MergedCells._cells,
                                                                      this._sourceRange._fromRow,
                                                                      this._sourceRange._fromCol,
                                                                      this._sourceRange._toRow,
                                                                      this._sourceRange._toCol);

        while (csem.Next())
        {
            if (!copiedMergedCells.ContainsKey(csem.Value))
            {
                ExcelAddress? adr = new ExcelAddress(worksheet.Name, worksheet.MergedCells._list[csem.Value]);
                eAddressCollition collideResult = this._sourceRange.Collide(adr);

                if (collideResult == eAddressCollition.Inside || collideResult == eAddressCollition.Equal)
                {
                    copiedMergedCells.Add(csem.Value,
                                          new ExcelAddress(this._destination._fromRow + (adr.Start.Row - this._sourceRange._fromRow),
                                                           this._destination._fromCol + (adr.Start.Column - this._sourceRange._fromCol),
                                                           this._destination._fromRow + (adr.End.Row - this._sourceRange._fromRow),
                                                           this._destination._fromCol + (adr.End.Column - this._sourceRange._fromCol)));
                }
                else
                {
                    //Partial merge of the address ignore.
                    copiedMergedCells.Add(csem.Value, null);
                }
            }
        }

        return copiedMergedCells;
    }

    private void CopyMergedCells(Dictionary<int, ExcelAddress> copiedMergedCells)
    {
        //Add merged cells
        foreach (ExcelAddress? m in copiedMergedCells.Values)
        {
            if (m != null)
            {
                this._destination._worksheet.MergedCells.Add(m, true);
            }
        }
    }

    private void CopyFullRow()
    {
        if (this._sourceRange._fromRow == 1 && this._sourceRange._toRow == ExcelPackage.MaxRows)
        {
            for (int col = 0; col < this._sourceRange.Columns; col++)
            {
                this._destination.Worksheet.Column(this._destination.Start.Column + col).OutlineLevel =
                    this._sourceRange.Worksheet.Column(this._sourceRange._fromCol + col).OutlineLevel;
            }
        }
    }

    private void CopyFullColumn()
    {
        if (this._sourceRange._fromCol == 1 && this._sourceRange._toCol == ExcelPackage.MaxColumns)
        {
            for (int row = 0; row < this._sourceRange.Rows; row++)
            {
                this._destination.Worksheet.Row(this._destination.Start.Row + row).OutlineLevel =
                    this._sourceRange.Worksheet.Row(this._sourceRange._fromRow + row).OutlineLevel;
            }
        }
    }
}