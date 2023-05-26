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
using System.Xml;
using System.Collections;
using System.Globalization;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Controls;
using System.Text;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Constants;
using System.IO;

namespace OfficeOpenXml.Drawing.Vml
{
    internal class ExcelVmlDrawingCollection
        : ExcelVmlDrawingBaseCollection, IEnumerable<ExcelVmlDrawingBase>, IDisposable, IPictureRelationDocument
    {
        internal CellStore<int> _drawingsCellStore;
        internal Dictionary<string, int> _drawingsDict = new Dictionary<string, int>();
        internal List<ExcelVmlDrawingBase> _drawings = new List<ExcelVmlDrawingBase>();
        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();
        internal ExcelVmlDrawingCollection(ExcelWorksheet ws, Uri uri) :
            base(ws, uri, "d:legacyDrawing/@r:id")
        {
            this._drawingsCellStore = new CellStore<int>();
            if (uri == null)
            {
                this.VmlDrawingXml.LoadXml(CreateVmlDrawings());
            }
            else
            {
                this.AddDrawingsFromXml(ws);
            }
        }
        ~ExcelVmlDrawingCollection()
        {
            this._drawingsCellStore?.Dispose();
            this._drawingsCellStore = null;
        }
        protected internal void AddDrawingsFromXml(ExcelWorksheet ws)
        {
            XmlNodeList? nodes = this.VmlDrawingXml.SelectNodes("//v:shape", this.NameSpaceManager);
            //var list = new List<IRangeID>();
            foreach (XmlNode node in nodes)
            {
                string? objectType = node.SelectSingleNode("x:ClientData/@ObjectType", this.NameSpaceManager)?.Value;
                ExcelVmlDrawingBase vmlDrawing;
                switch (objectType)
                {
                    case "Drop":
                    case "List":
                    case "Button":
                    case "GBox":
                    case "Label":
                    case "Checkbox":
                    case "Spin":
                    case "Radio":
                    case "EditBox":
                    case "Dialog":
                        vmlDrawing = new ExcelVmlDrawingControl(this._ws, node, this.NameSpaceManager);
                        this._drawings.Add(vmlDrawing);
                        break;
                    default:    //Comments
                        XmlNode? rowNode = node.SelectSingleNode("x:ClientData/x:Row", this.NameSpaceManager);
                        XmlNode? colNode = node.SelectSingleNode("x:ClientData/x:Column", this.NameSpaceManager);
                        int row, col;
                        if (rowNode != null && colNode != null)
                        {
                            row = int.Parse(rowNode.InnerText) + 1;
                            col = int.Parse(colNode.InnerText) + 1;
                        }
                        else
                        {
                            row = 1;
                            col = 1;
                        }
                        vmlDrawing = new ExcelVmlDrawingComment(node, ws.Cells[row, col], this.NameSpaceManager);
                        this._drawings.Add(vmlDrawing);
                        this._drawingsCellStore.SetValue(row, col, this._drawings.Count-1);
                        break;
                }
                string? id = string.IsNullOrEmpty(vmlDrawing.SpId) ? vmlDrawing.Id : vmlDrawing.SpId;
                int x = 2;
                if(this._drawingsDict.ContainsKey(id))
                {
                    while(this._drawingsDict.ContainsKey($"{id}-{x}"))
                    {
                        x++;
                    }
                    id = $"{id}-{x}";
                    if (string.IsNullOrEmpty(vmlDrawing.SpId))
                    {
                        vmlDrawing.Id= id;
                    }
                    else
                    {
                        vmlDrawing.SpId = id;
                    }
                }

                this._drawingsDict.Add(id, this._drawings.Count - 1);
            }
        }

        private static string CreateVmlDrawings()
        {
            string vml = string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">",
                ExcelPackage.schemaMicrosoftVml,
                ExcelPackage.schemaMicrosoftOffice,
                ExcelPackage.schemaMicrosoftExcel);

            vml += "<o:shapelayout v:ext=\"edit\">";
            vml += "<o:idmap v:ext=\"edit\" data=\"1\"/>";
            vml += "</o:shapelayout>";
            vml += "<v:shapetype path=\"m,l,21600r21600,l21600,xe\" o:spt=\"201\" coordsize=\"21600,21600\" id=\"_x0000_t201\">"; 
            vml += "<v:stroke joinstyle=\"miter\"/>";
            vml += "<v:path o:connecttype=\"rect\" fillok=\"f\" strokeok=\"f\" o:extrusionok=\"f\" shadowok=\"f\"/>";
            vml += "<o:lock v:ext=\"edit\" shapetype=\"t\"/>";
            vml += "</v:shapetype>";
            vml += "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">";
            vml += "<v:stroke joinstyle=\"miter\" />";
            vml += "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />";
            vml += "</v:shapetype>";
            vml += "</xml>";

            return vml;
        }
        internal ExcelVmlDrawingComment AddComment(ExcelRangeBase cell)
        {
            XmlNode node = this.AddCommentDrawing(cell);
            ExcelVmlDrawingComment? draw = new ExcelVmlDrawingComment(node, cell, this.NameSpaceManager);
            this._drawings.Add(draw);
            this._drawingsCellStore.SetValue(cell._fromRow, cell._fromCol, this._drawings.Count-1);
            return draw;
        }
        private XmlNode AddCommentDrawing(ExcelRangeBase cell)
        {
            this.CreateVmlPart(); //Create the vml part to be able to create related parts (like blip fill images).
            int row = cell.Start.Row, col = cell.Start.Column;
            XmlElement? node = this.VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);

            int r = cell._fromRow, c = cell._fromCol;
            bool prev = this._drawingsCellStore.PrevCell(ref r, ref c);
            if (prev)
            {                
                ExcelVmlDrawingBase? prevDraw = this._drawings[this._drawingsCellStore.GetValue(r, c)];
                prevDraw.TopNode.ParentNode.InsertBefore(node, prevDraw.TopNode);
            }
            else
            {
                this.VmlDrawingXml.DocumentElement.AppendChild(node);
            }

            node.SetAttribute("id", this.GetNewId());
            node.SetAttribute("type", "#_x0000_t202");
            node.SetAttribute("style", "position:absolute;z-index:1; visibility:hidden");
            //node.SetAttribute("style", "position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1; visibility:hidden"); 
            node.SetAttribute("fillcolor", "#ffffe1");
            node.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");

            string vml = "<v:fill color2=\"#ffffe1\" />";
            vml += "<v:shadow on=\"t\" color=\"black\" obscured=\"t\" />";
            vml += "<v:path o:connecttype=\"none\" />";
            vml += "<v:textbox style=\"mso-direction-alt:auto\">";
            vml += "<div style=\"text-align:left\" />";
            vml += "</v:textbox>";
            vml += "<x:ClientData ObjectType=\"Note\">";
            vml += "<x:MoveWithCells />";
            vml += "<x:SizeWithCells />";
            vml += string.Format("<x:Anchor>{0}, 15, {1}, 2, {2}, 31, {3}, 1</x:Anchor>", col, row - 1, col + 2, row + 3);
            vml += "<x:AutoFill>False</x:AutoFill>";
            vml += string.Format("<x:Row>{0}</x:Row>", row - 1);
            vml += string.Format("<x:Column>{0}</x:Column>", col - 1);
            vml += "</x:ClientData>";

            node.InnerXml = vml;
            return node;
        }
        internal ExcelVmlDrawingControl AddControl(ExcelControl ctrl, string name)
        {
            XmlNode node = this.AddControlDrawing(ctrl, name);
            ExcelVmlDrawingControl? draw = new ExcelVmlDrawingControl(this._ws, node, this.NameSpaceManager);
            this._drawings.Add(draw);
            if(this._drawingsDict.ContainsKey(draw.Id) == false)
            {
                this._drawingsDict.Add(draw.Id, this._drawings.Count - 1);
            }
            return draw;
        }
        private XmlNode AddControlDrawing(ExcelControl ctrl, string name)
        {
            this.CreateVmlPart(); //Create the vml part to be able to create related parts (like blip fill images).
            XmlElement? shapeElement = this.VmlDrawingXml.CreateElement("v", "shape", ExcelPackage.schemaMicrosoftVml);

            this.VmlDrawingXml.DocumentElement.AppendChild(shapeElement);

            shapeElement.SetAttribute("spid", ExcelPackage.schemaMicrosoftOffice, "_x0000_s"+ctrl.Id);
            shapeElement.SetAttribute("id", name);
            //shapeElement.SetAttribute("id", $"{ctrl.ControlTypeString}_x{ctrl.Id}_1");
            shapeElement.SetAttribute("type", "#_x0000_t201");
            shapeElement.SetAttribute("style", "position:absolute;z-index:1;");
            shapeElement.SetAttribute("insetmode", ExcelPackage.schemaMicrosoftOffice, "auto");
            SetShapeAttributes(ctrl, shapeElement);

            StringBuilder? vml = new StringBuilder();
            vml.Append(GetVml(ctrl, shapeElement));
            vml.Append("<o:lock v:ext=\"edit\" rotation=\"t\"/>");
            vml.Append("<v:textbox style=\"mso-direction-alt:auto\" o:singleclick=\"f\">");
            if (ctrl is ExcelControlWithText textControl)
            {
                vml.Append($"<div style=\"text-align:center\"><font color=\"#000000\" size=\"{GetFontSize(ctrl)}\" face=\"{GetFontName(ctrl)}\">{textControl.Text}</font></div>");
            }
            vml.Append("</v:textbox>");
            vml.Append($"<x:ClientData ObjectType=\"{ctrl.ControlTypeString}\">");
            vml.Append(string.Format("<x:Anchor>{0}</x:Anchor>", ctrl.GetVmlAnchorValue()));
            vml.Append(GetVmlClientData(ctrl, shapeElement));
            vml.Append("<x:PrintObject>False</x:PrintObject>");
            vml.Append("<x:AutoFill>False</x:AutoFill>");
            if (ctrl.ControlType != eControlType.GroupBox)
            {
                vml.Append("<x:TextVAlign>Center</x:TextVAlign>");
            }

            vml.Append("</x:ClientData>");

            shapeElement.InnerXml = vml.ToString();
            return shapeElement;
        }
        private static string GetFontName(ExcelControl ctrl)
        {
            if (ctrl.ControlType == eControlType.Button)
            {
                return "Calibri";
            }
            else
            {
                return "Segoe UI";
            }
        }

        private static string GetFontSize(ExcelControl ctrl)
        {
            if (ctrl.ControlType == eControlType.Button)
            {
                return "220";
            }
            else
            {
                return "160";
            }
        }

        private static string GetVmlClientData(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    return "<x:TextHAlign>Center</x:TextHAlign>";
                case eControlType.CheckBox:
                case eControlType.GroupBox:
                    return "<x:SizeWithCells/><x:NoThreeD/>";
                case eControlType.RadioButton:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:NoThreeD/><x:FirstButton/>";
                case eControlType.DropDown:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>1</x:Page><x:Dx>22</x:Dx><x:Sel>0</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT><x:DropStyle>Combo</x:DropStyle>   <x:DropLines>8</x:DropLines>";
                case eControlType.ListBox:
                    return "<x:SizeWithCells/><x:AutoLine>False</x:AutoLine><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>7</x:Page><x:Dx>22</x:Dx><x:Sel>0</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT>";
                case eControlType.Label:
                    return "<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>";
                case eControlType.ScrollBar:
                    return "<x:SizeWithCells/><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>100</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page><x:Dx>22</x:Dx>";
                case eControlType.SpinButton:
                    return "   <x:Val>0</x:Val><x:Min>0</x:Min><x:Max>30000</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page><x:Dx>22</x:Dx>";
                default:
                    return "";
            }
        }

        private static string GetVml(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    return "<v:fill o:detectmouseclick=\"t\" color2=\"buttonFace[67]\"/>";
                case eControlType.CheckBox:
                    return "<v:path fillok=\"t\" strokeok=\"t\" shadowok=\"t\"/>";
                default:
                    return "";
            }
        }

        private static void SetShapeAttributes(ExcelControl ctrl, XmlElement shapeElement)
        {
            switch (ctrl.ControlType)
            {
                case eControlType.Button:
                    shapeElement.SetAttribute("fillcolor", "buttonFace [67]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("button", ExcelPackage.schemaMicrosoftOffice, "t");
                    break;
                case eControlType.CheckBox:
                case eControlType.RadioButton:
                    shapeElement.SetAttribute("fillcolor", "windows [65]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    //shapeElement.SetAttribute("button", ExcelPackage.schemaMicrosoftOffice, "t");
                    shapeElement.SetAttribute("stroked", "f");
                    shapeElement.SetAttribute("filled", "f");
                    //style = "position:absolute; margin-left:15pt;margin-top:10.5pt;width:120.75pt;height:23.25pt;z-index:1; mso-wrap-style:tight" type = "#_x0000_t201" >
                    break;
                case eControlType.ListBox:
                case eControlType.DropDown:
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("stroked", "f");
                    break;
                case eControlType.ScrollBar:
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    break;
                case eControlType.Label:
                    shapeElement.SetAttribute("fillcolor", "windows [65]");
                    shapeElement.SetAttribute("strokecolor", "windowText [64]");
                    shapeElement.SetAttribute("stroked", "f");
                    shapeElement.SetAttribute("filled", "f");
                    break;
            }

        }

        int _nextID = 0;
        /// <summary>
        /// returns the next drawing id.
        /// </summary>
        /// <returns></returns>
        internal string GetNewId()
        {
            if (this._nextID == 0)
            {
                foreach (ExcelVmlDrawingBase draw in this)
                {
                    if (draw.Id.Length > 3 && draw.Id.StartsWith("vml", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(draw.Id.Substring(3, draw.Id.Length - 3), NumberStyles.Any, CultureInfo.InvariantCulture, out int id))
                        {
                            if (id > this._nextID)
                            {
                                this._nextID = id;
                            }
                        }
                    }
                }
            }

            this._nextID++;
            return "vml" + this._nextID.ToString();
        }
        internal ExcelVmlDrawingBase this[string id]
        {
            get
            {
                if(this._drawingsDict.ContainsKey(id))
                {
                    return this._drawings[this._drawingsDict[id]];
                }
                return null;
            }
        }

        internal ExcelVmlDrawingBase this[int row, int column]
        {
            get
            {
                return this._drawings[this._drawingsCellStore.GetValue(row, column)];
            }
        }
        internal bool ContainsKey(int row, int column)
        {
            return this._drawingsCellStore.Exists(row, column);
        }
        internal int Count
        {
            get
            {
                return this._drawings.Count;
            }
        }

        public ExcelPackage Package => this._package;

        public Dictionary<string, HashInfo> Hashes => this._hashes;

        public ZipPackagePart RelatedPart => this.Part;

        public Uri RelatedUri => this.Uri;
        #region "Enumerator"
        //CellStoreEnumerator<ExcelVmlDrawingComment> _enum;
        public IEnumerator<ExcelVmlDrawingBase> GetEnumerator()
        {
            //Reset();
            return this._drawings.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            //Reset();
            return this._drawings.GetEnumerator();
        }

        ///// <summary>
        ///// The current range when enumerating
        ///// </summary>
        //public ExcelVmlDrawingComment Current
        //{
        //    get
        //    {
        //        return _enum.Current;
        //    }
        //}

        ///// <summary>
        ///// The current range when enumerating
        ///// </summary>
        //object IEnumerator.Current
        //{
        //    get
        //    {
        //        return _enum.Current;
        //    }
        //}

        //public bool MoveNext()
        //{
        //    return _enum.Next();
        //}

        //public void Reset()
        //{
        //    if (_enum != null) _enum.Dispose();
        //     _enum = new CellStoreEnumerator<ExcelVmlDrawingComment>(_drawingsCellStore, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        //}
        void IDisposable.Dispose()
        {
            this._drawingsCellStore.Dispose();
        }

        //public void Dispose()
        //{
        //    throw new NotImplementedException();
        //}
        #endregion
    } 
}
