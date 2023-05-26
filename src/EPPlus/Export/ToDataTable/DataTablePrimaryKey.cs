/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTablePrimaryKey
    {
        private readonly ToDataTableOptions _options;
        private readonly HashSet<string> _keyNames = new HashSet<string>();

        public DataTablePrimaryKey(ToDataTableOptions options)
        {
            this._options = options;
            this.Initialize();
        }

        private void Initialize()
        {
            if(this._options.PrimaryKeyNames.Any())
            {
                foreach(string? name in this._options.PrimaryKeyNames)
                {
                    this.AddPrimaryKeyName(name);
                }
            }
            else if(this._options.PrimaryKeyIndexes.Any())
            {
                foreach(int ix in this._options.PrimaryKeyIndexes)
                {
                    try
                    {
                        DataColumnMapping? mapping = this._options.Mappings.GetByRangeIndex(ix);
                        this.AddPrimaryKeyName(mapping.DataColumnName);
                    }
                    catch(ArgumentOutOfRangeException e)
                    {
                        throw new ArgumentOutOfRangeException("primary key index out of range: " + ix, e);
                    }
                }
            }
        }

        private void AddPrimaryKeyName(string name)
        {
            if (this._keyNames.Contains(name))
            {
                throw new InvalidOperationException("Duplicate primary key name: " + name);
            }
            if (!this._options.Mappings.Exists(x => x.DataColumnName == name))
            {
                throw new InvalidOperationException("Invalid primary key name, no corresponding DataColumn: " + name);
            }

            this._keyNames.Add(name);
        }

        internal IEnumerable<string> KeyNames => this._keyNames;

        internal bool HasKeys => this._keyNames.Any();

        internal bool ContainsKey(string key)
        {
            return this._keyNames.Contains(key);
        }
    }
}
