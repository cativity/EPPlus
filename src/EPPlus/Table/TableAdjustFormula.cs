using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table
{
    internal class TableAdjustFormula
    {
        ExcelTable _tbl;
        public TableAdjustFormula(ExcelTable tbl)
        {
            _tbl = tbl;
        }

        internal void AdjustFormulas(string prevName, string name)
        {
            foreach (ExcelWorksheet? ws in _tbl.WorkSheet.Workbook.Worksheets)
            {
                foreach (ExcelTable? tbl in ws.Tables)
                {
                    foreach (ExcelTableColumn? c in tbl.Columns)
                    {
                        if (!string.IsNullOrEmpty(c.CalculatedColumnFormula))
                        {
                            c.CalculatedColumnFormula = ReplaceTableName(c.CalculatedColumnFormula, prevName, name);
                        }
                    }
                }

                CellStoreEnumerator<object>? cse = new CellStoreEnumerator<object>(ws._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string f)
                    {
                        if (f.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                        {
                            ws._formulas.SetValue(cse.Row, cse.Column, ReplaceTableName(f, prevName, name));
                        }
                    }
                }

                foreach (ExcelWorksheet.Formulas? sf in ws._sharedFormulas.Values)
                {
                    if (sf.Formula.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        sf.Formula = ReplaceTableName(sf.Formula, prevName, name);
                    }
                }

                foreach (ExcelNamedRange? n in ws.Names)
                {
                    AdjustName(n, prevName, name);
                }
            }

            foreach (ExcelNamedRange? n in _tbl.WorkSheet.Workbook.Names)
            {
                AdjustName(n, prevName, name);
            }
        }

        private void AdjustName(ExcelNamedRange n, string prevName, string name)
        {
            if (!string.IsNullOrEmpty(n.Formula))
            {
                if (n.Formula.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    n.Formula = ReplaceTableName(n.Formula, prevName, name);
                }
            }
            else if (n.IsName == false)
            {
                if (n.Address.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    n.Address = ReplaceTableName(n.Address, prevName, name);
                }
            }
        }

        private string ReplaceTableName(string formula, string prevName, string name)
        {
            IEnumerable<Token>? tokens = _tbl.WorkSheet.Workbook.FormulaParser.Lexer.Tokenize(formula);
            string? f = "";
            foreach (Token t in tokens)
            {
                if (t.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    ExcelAddressBase? a = new ExcelAddressBase(t.Value);
                    if (a.Table == null)
                    {
                        f += t.Value;
                    }
                    else
                    {
                        f += a.ChangeTableName(prevName, name);
                    }
                }
                else
                {
                    f += t.Value;
                }
            }

            return f;
        }
    }
}
