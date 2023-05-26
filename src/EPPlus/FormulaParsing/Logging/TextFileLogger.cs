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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace OfficeOpenXml.FormulaParsing.Logging
{
    internal class TextFileLogger : IFormulaParserLogger
    {
        private StreamWriter _sw;
        private const string Separator = "=================================";
        private int _count;
        private DateTime _startTime = DateTime.Now;
        private Dictionary<string, int> _funcs = new Dictionary<string, int>();
        private Dictionary<string, long> _funcPerformance = new Dictionary<string, long>();
        internal TextFileLogger(FileInfo fileInfo)
        {
#if (Core)
            this._sw = new StreamWriter(new FileStream(fileInfo.FullName, FileMode.Append));
#else
            _sw = new StreamWriter(fileInfo.FullName);
#endif
        }

        private void WriteSeparatorAndTimeStamp()
        {
            this._sw.WriteLine(Separator);
            this._sw.WriteLine("Timestamp: {0}", DateTime.Now);
            this._sw.WriteLine();
        }

        private void WriteAddressInfo(ParsingContext context)
        {
            if (context.Scopes.Current != null && context.Scopes.Current.Address != null)
            {
                this._sw.WriteLine("Worksheet: {0}", context.Scopes.Current.Address.Worksheet ?? "<not specified>");
                this._sw.WriteLine("Address: {0}", context.Scopes.Current.Address.Address ?? "<not available>");
            }
        }

        public void Log(ParsingContext context, Exception ex)
        {
            this.WriteSeparatorAndTimeStamp();
            this.WriteAddressInfo(context);
            this._sw.WriteLine(ex);
            this._sw.WriteLine();
        }

        public void Log(ParsingContext context, string message)
        {
            this.WriteSeparatorAndTimeStamp();
            this.WriteAddressInfo(context);
            this._sw.WriteLine(message);
            this._sw.WriteLine();
        }

        public void Log(string message)
        {
            this.WriteSeparatorAndTimeStamp();
            this._sw.WriteLine(message);
            this._sw.WriteLine();
        }

        public void LogCellCounted()
        {
            this._count++;
            if (this._count%500 == 0)
            {
                this._sw.WriteLine(Separator);
                TimeSpan timeEllapsed = DateTime.Now.Subtract(this._startTime);
                this._sw.WriteLine("{0} cells parsed, time {1} seconds", this._count, timeEllapsed.TotalSeconds);

                List<string>? funcs = this._funcs.Keys.OrderByDescending(x => this._funcs[x]).ToList();
                foreach (string? func in funcs)
                {
                    this._sw.Write(func + "  - " + this._funcs[func]);
                    if (this._funcPerformance.ContainsKey(func))
                    {
                        this._sw.Write(" - avg: " + this._funcPerformance[func]/ this._funcs[func] + " milliseconds");
                    }

                    this._sw.WriteLine();
                }

                this._sw.WriteLine();
                this._funcs.Clear();

            }
        }

        public void LogFunction(string func)
        {
            if (!this._funcs.ContainsKey(func))
            {
                this._funcs.Add(func, 0);
            }

            this._funcs[func]++;
        }

        public void LogFunction(string func, long milliseconds)
        {
            if (!this._funcPerformance.ContainsKey(func))
            {
                this._funcPerformance[func] = 0;
            }

            this._funcPerformance[func] += milliseconds;
        }

        public void Dispose()
        {
            this._sw.Close();
            this._sw.Dispose();
        }
    }
}
