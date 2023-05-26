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
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// A collection of ignored errors per range for a worksheet
    /// </summary>
    public class ExcelIgnoredErrorCollection : IEnumerable<ExcelIgnoredError>, IDisposable
    {
        private ExcelPackage _package;
        private ExcelWorksheet _excelWorksheet;
        private XmlNamespaceManager _nameSpaceManager;
        private List<ExcelIgnoredError> _list = new List<ExcelIgnoredError>();
        internal ExcelIgnoredErrorCollection(ExcelPackage package, ExcelWorksheet excelWorksheet, XmlNamespaceManager nameSpaceManager)
        {
            this._package = package;
            this._excelWorksheet = excelWorksheet;
            this._nameSpaceManager = nameSpaceManager;
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="Index">This index</param>
        /// <returns></returns>
        public ExcelIgnoredError this[int Index]
        {
            get
            {
                if(Index<0 || Index> this._list.Count)
                {
                    throw new ArgumentOutOfRangeException("Index");
                }
                return this._list[Index];
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return this._list.Count;
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._list.GetEnumerator();
        }
        /// <summary>
        /// Adds an IgnoreError item to the collection
        /// </summary>
        /// <param name="address">The address to add</param>
        /// <returns>The IgnoreError Item</returns>
        public ExcelIgnoredError Add(ExcelAddressBase address)
        {

            XmlElement? node = this._excelWorksheet.WorksheetXml.CreateElement("ignoredError", ExcelPackage.schemaMain);
            this.TopNode.AppendChild(node);
            ExcelIgnoredError? item = new ExcelIgnoredError(this._nameSpaceManager, node, address);
            this._list.Add(item);
            return item;
        }
        XmlNode _topNode=null;
        internal XmlNode TopNode
        {
            get { return this._topNode ??= this._excelWorksheet.CreateNode("d:ignoredErrors"); }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator<ExcelIgnoredError> IEnumerable<ExcelIgnoredError>.GetEnumerator()
        {
            return this._list.GetEnumerator();
        }
        /// <summary>
        /// Called when the class is disposed.
        /// </summary>
        public void Dispose()
        {
            
        }
    }
}