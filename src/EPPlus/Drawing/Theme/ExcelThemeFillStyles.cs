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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// Defines fill styles for a theme.
    /// </summary>
    public class ExcelThemeFillStyles : XmlHelper, IEnumerable<ExcelDrawingFill>
    {
        private readonly List<ExcelDrawingFill> _list;
        internal ExcelThemeFillStyles(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelThemeBase theme) : base(nameSpaceManager, topNode)
        {
            this._list = new List<ExcelDrawingFill>();
            foreach (XmlNode node in topNode.ChildNodes)
            {
                this._list.Add(new ExcelDrawingFill(theme, nameSpaceManager, node, "", this.SchemaNodeOrder));
            }
        }
        /// <summary>
        /// Get the enumerator for the Theme
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelDrawingFill> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._list.GetEnumerator();
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>The fill</returns>
        public ExcelDrawingFill this[int index]
        {
            get
            {
                return this._list[index];
            }
        }
        /// <summary>
        /// Adds a new fill to the collection
        /// </summary>
        /// <param name="style">The fill style</param>
        /// <returns>The fill</returns>
        public ExcelDrawingFill Add(eFillStyle style)
        {            
            XmlElement? node = this.TopNode.OwnerDocument.CreateElement("a",ExcelDrawingFillBasic.GetStyleText(style),  ExcelPackage.schemaMain);
            this.TopNode.AppendChild(node);
            return new ExcelDrawingFill(null, this.NameSpaceManager, this.TopNode, "", this.SchemaNodeOrder);
        }
        /// <summary>
        /// Remove a fill item
        /// </summary>
        /// <param name="item">The item</param>
        public void Remove(ExcelDrawingFill item)
        {
            if(this._list.Count==3)
            {
                throw (new InvalidOperationException("Collection must contain at least 3 items"));
            }

            if (this._list.Contains(item))
            {
                this._list.Remove(item);
                item.TopNode.ParentNode.RemoveChild(item.TopNode);
            }
        }
        /// <summary>
        /// Remove the item at the specified index
        /// </summary>
        /// <param name="Index"></param>
        public void Remove(int Index)
        {
            if(Index >= this._list.Count)
            {
                throw new ArgumentException("Index", "Index out of range");
            }

            this._list.Remove(this._list[Index]);            
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
    }
}
