/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// An abstract class inherited by form controls
/// </summary>
public abstract class ExcelControl : ExcelDrawing
{
    internal ExcelVmlDrawingControl _vml;
    internal XmlHelper _ctrlProp;
    internal XmlHelper _vmlProp;
    internal ControlInternal _control;

    internal ExcelControl(ExcelDrawings drawings, XmlNode drawingNode, ControlInternal control, ZipPackagePart ctrlPropPart, XmlDocument ctrlPropXml, ExcelGroupShape parent = null) :
        base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
    {
        this._control = control;
        this._vml = (ExcelVmlDrawingControl)drawings.Worksheet.VmlDrawings[this.LegacySpId];
        this._vmlProp = XmlHelperFactory.Create(this._vml.NameSpaceManager, this._vml.GetNode("x:ClientData"));
        this.ControlPropertiesXml = ctrlPropXml;
        this.ControlPropertiesPart = ctrlPropPart;
        this.ControlPropertiesUri = ctrlPropPart.Uri;
        this._ctrlProp = XmlHelperFactory.Create(this.NameSpaceManager, ctrlPropXml.DocumentElement);
    }
    internal ExcelControl(ExcelDrawings drawings, XmlNode drawingNode, string name, ExcelGroupShape parent = null) : 
        base(drawings, drawingNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
    {
        ExcelWorksheet? ws = drawings.Worksheet;
                       
        //Drawing Xml
        XmlElement spElement = this.CreateShapeNode();
        spElement.InnerXml = this.ControlStartDrawingXml();
        this.CreateClientData();

        this.ControlPropertiesXml = new XmlDocument();
        this.ControlPropertiesXml.LoadXml(this.ControlStartControlPrXml());            
        int id= ws.SheetId;
        this.ControlPropertiesUri = GetNewUri(ws._package.ZipPackage, "/xl/ctrlProps/ctrlProp{0}.xml",ref id);
        this.ControlPropertiesPart = ws._package.ZipPackage.CreatePart(this.ControlPropertiesUri, ContentTypes.contentTypeControlProperties);
        ZipPackageRelationship? rel=ws.Part.CreateRelationship(this.ControlPropertiesUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/ctrlProp");

        //Vml
        this._vml=drawings.Worksheet.VmlDrawings.AddControl(this, name);
        this._vmlProp = XmlHelperFactory.Create(this._vml.NameSpaceManager, this._vml.GetNode("x:ClientData"));

        //Control in worksheet xml
        XmlNode ctrlNode = ws.CreateControlContainerNode();
        ((XmlElement)ws.TopNode).SetAttribute("xmlns:xdr", ExcelPackage.schemaSheetDrawings);   //Make sure the namespace exists
        ((XmlElement)ws.TopNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);   //Make sure the namespace exists
        ((XmlElement)ws.TopNode).SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);   //Make sure the namespace exists
        ctrlNode.InnerXml = this.GetControlStartWorksheetXml(rel.Id);
        this._control = new ControlInternal(this.NameSpaceManager, ctrlNode.FirstChild);
        this._ctrlProp = XmlHelperFactory.Create(this.NameSpaceManager, this.ControlPropertiesXml.DocumentElement);
    }
    private string GetControlStartWorksheetXml(string relId)
    {
        StringBuilder? sb = new StringBuilder();

        sb.Append($"<control shapeId=\"{this.Id}\" r:id=\"{relId}\" name=\"\">");
        sb.Append("<controlPr defaultSize=\"0\" print=\"0\" autoFill=\"0\" autoPict=\"0\">");
        if (this.ControlType == eControlType.Label)
        {
            sb.Append("<anchor moveWithCells=\"1\" sizeWithCells=\"1\">");
        }
        else if(this.ControlType == eControlType.Button)
        {
            sb.Append("<anchor>");
        }
        else
        {
            sb.Append("<anchor moveWithCells=\"1\" >");
        }
        sb.Append($"<from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>");
        sb.Append($"<to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></to>");
        sb.Append("</anchor></controlPr></control>");
        return sb.ToString();
    }
    private string ControlStartControlPrXml()
    {
        string? xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" {0} />";
        switch (this.ControlType)
        {
            case eControlType.Button:
                return string.Format(xml, "objectType=\"Button\" lockText=\"1\"");
            case eControlType.CheckBox:
                return string.Format(xml, "objectType=\"CheckBox\" lockText=\"1\" noThreeD=\"1\"");
            case eControlType.RadioButton:
                return string.Format(xml, "objectType=\"Radio\" lockText=\"1\" noThreeD=\"1\"");
            case eControlType.DropDown:
                return string.Format(xml, "objectType=\"Drop\" dropStyle=\"combo\" dx=\"22\" noThreeD=\"1\" sel=\"0\" val=\"0\"");
            case eControlType.ListBox:
                return string.Format(xml, "objectType=\"List\" dx=\"22\" noThreeD=\"1\" sel=\"0\" val=\"0\"");
            case eControlType.Label:
                return string.Format(xml, "objectType=\"Label\" lockText=\"1\"");
            case eControlType.ScrollBar:
                return string.Format(xml, "objectType=\"Scroll\" dx=\"22\" max=\"100\" page=\"10\" val=\"0\"");
            case eControlType.SpinButton:
                return string.Format(xml, "objectType=\"Spin\" dx=\"22\" max=\"30000\" page=\"10\" val=\"0\"");
            case eControlType.GroupBox:
                return string.Format(xml, "objectType=\"GBox\" noThreeD=\"1\"");
            default:
                throw new NotImplementedException();
        }
    }

    private string ControlStartDrawingXml()
    {
        StringBuilder xml = new StringBuilder();
        xml.Append($"<xdr:nvSpPr><xdr:cNvPr hidden=\"1\" name=\"\" id=\"{this._id}\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{this._id}\"/></a:ext><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId id=\"{{00000000-0008-0000-0000-000001040000}}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>");
        xml.Append($"<xdr:spPr bwMode=\"auto\"><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        switch (this.ControlType)
        {
            case eControlType.Button:
                xml.Append($"<a:noFill/><a:ln w=\"9525\"><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln>");
                break;
            case eControlType.CheckBox:
            case eControlType.RadioButton:
                xml.Append($"<a:noFill/><a:ln><a:noFill/></a:ln><a:extLst><a:ext uri=\"{{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}}\"><a14:hiddenFill><a:solidFill><a:srgbClr val=\"FFFFFF\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"65\"/></a:solidFill></a14:hiddenFill></a:ext><a:ext uri=\"{{91240B29-F687-4F45-9708-019B960494DF}}\"><a14:hiddenLine w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"64\"/></a:solidFill><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a14:hiddenLine></a:ext></a:extLst>");
                break;
            case eControlType.ListBox:
                xml.Append("<a:noFill/><a:ln><a:noFill/></a:ln><a:extLst><a:ext uri=\"{{91240B29-F687-4F45-9708-019B960494DF}}\"><a14:hiddenLine w=\"9525\"><a:noFill/><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a14:hiddenLine></a:ext></a:extLst>");
                break;
        }
        xml.Append("</xdr:spPr>");
        if(this is ExcelControlWithText)
        {
            xml.Append($"<xdr:txBody><a:bodyPr upright=\"1\" anchor=\"{this.GetDrawingAnchor()}\" bIns=\"27432\" rIns=\"27432\" tIns=\"27432\" lIns=\"27432\" wrap=\"square\" vertOverflow=\"clip\"/>" +
                       $"<a:lstStyle/>" +
                       $"<a:p>{GetrPr(this.ControlType)}" +
                       $"<a:r><a:rPr lang=\"en-US\" sz=\"{this.GetFontSize()}\" baseline=\"0\" strike=\"noStrike\" u=\"none\" i=\"0\" b=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"{this.GetFontName()}\"/><a:cs typeface=\"{this.GetFontName()}\"/></a:rPr><a:t></a:t></a:r></a:p></xdr:txBody>");
        }
        return xml.ToString();
    }

    private string GetFontName()
    {
        if (this.ControlType == eControlType.Button)
        {
            return "Calibri";
        }
        else
        {
            return "Segoe UI";
        }
    }

    private string GetFontSize()
    {
        if(this.ControlType==eControlType.Button)
        {
            return "1100";
        }
        else
        {
            return "800";
        }
    }

    private string GetDrawingAnchor()
    {
        if(this.ControlType==eControlType.GroupBox)
        {
            return "t";
        }
        else
        {
            return "ctr";
        }
    }

    private static string GetrPr(eControlType controlType)
    {
        switch (controlType)
        {
            case eControlType.Button:
                return "<a:pPr rtl=\"0\" algn=\"ctr\"><a:defRPr sz=\"1000\"/></a:pPr>";
            case eControlType.CheckBox:
            case eControlType.RadioButton:
            case eControlType.Label:
                return "<a:pPr rtl=\"0\" algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>"; 
            default:
                return "<a:pPr rtl=\"0\" algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>";                    
        }
    }

    private XmlNode GetVmlNode(ExcelVmlDrawingCollection vmlDrawings)
    {
        return vmlDrawings.FirstOrDefault(x => x.Id == this.LegacySpId)?.TopNode;
    }

    /// <summary>
    /// The control property xml associated with the control
    /// </summary>
    public XmlDocument ControlPropertiesXml { get; private set; }
    internal ZipPackagePart ControlPropertiesPart { get; private set; }
    internal Uri ControlPropertiesUri { get; private set; }
    /// <summary>
    /// The type of form control
    /// </summary>
    public abstract eControlType ControlType
    {
        get;
    }
    internal string ControlTypeString
    {
        get
        {
            switch(this.ControlType)
            {
                case eControlType.GroupBox:
                    return "GBox";
                case eControlType.CheckBox:
                    return "Checkbox";
                case eControlType.RadioButton:
                    return "Radio";
                case eControlType.DropDown:
                    return "Drop";
                case eControlType.ListBox:
                    return "List";
                case eControlType.SpinButton:
                    return "Spin";
                case eControlType.ScrollBar:
                    return "Scroll";
                default:
                    return this.ControlType.ToString();
            }
        }
    }
    internal string LegacySpId
    {
        get
        {
            return this.GetXmlNodeString($"{this.GetlegacySpIdPath()}/a:extLst/a:ext[@uri='{ExtLstUris.LegacyObjectWrapperUri}']/a14:compatExt/@spid");
        }
        set
        {
            XmlNode? node = this.GetNode(this.GetlegacySpIdPath());
            XmlHelper? extHelper = XmlHelperFactory.Create(this.NameSpaceManager, node);
            XmlNode? extNode= extHelper.GetOrCreateExtLstSubNode(ExtLstUris.LegacyObjectWrapperUri, "a14");
            if (extNode.InnerXml == "")
            {
                extNode.InnerXml = $"<a14:compatExt/>";
            }
            ((XmlElement)extNode.FirstChild).SetAttribute("spid", value);
        }
    }
    internal string GetlegacySpIdPath()
    {
        return $"{(this._topPath == "" ? "" : this._topPath + "/")}xdr:nvSpPr/xdr:cNvPr";
    }

    /// <summary>
    /// The name of the control
    /// </summary>
    public override string Name
    {
        get
        {
            return this._control.Name;
        }
        set
        {
            this._control.Name=value;
            this._vml.Id = value;
            base.Name = value;
        }
    }
    internal eEditAs GetCellAnchorFromWorksheetXml()
    {
        if (this._control.MoveWithCells && this._control.SizeWithCells)
        {
            return eEditAs.TwoCell;
        }
        else if(this._control.MoveWithCells)
        {
            return eEditAs.OneCell;
        }
        else
        {
            return eEditAs.Absolute;
        }
    }
    internal void SetCellAnchor(eEditAs value)
    {
        switch(value)
        {
            case eEditAs.Absolute:
                this._control.MoveWithCells = false;
                this._control.SizeWithCells = false;
                break;
            case eEditAs.OneCell:
                this._control.MoveWithCells = true;
                this._control.SizeWithCells = false;
                break;
            default:
                this._control.MoveWithCells = true;
                this._control.SizeWithCells = true;
                break;
        }
    }

    /// <summary>
    /// Gets or sets the alternative text for the control.
    /// </summary>
    public string AlternativeText
    {
        get
        {
            return this._control.AlternativeText;
        }
        set
        {
            this._control.AlternativeText = value;
        }
    }
    /// <summary>
    /// Gets or sets the macro function assigned.
    /// </summary>
    public string Macro
    {
        get
        {
            return this._control.Macro;
        }
        set
        {
            this._control.Macro = value;
            this._vmlProp.SetXmlNodeString("x:FmlaMacro", value);
        }
    }

    internal string GetVmlAnchorValue()
    {
        ExcelPosition? from = this._control?.From ?? this.From;
        ExcelPosition? to = this._control?.To ?? this.To;
        return string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}",
                             from.Column, Math.Round(from.ColumnOff / EMU_PER_PIXEL * 1.5),
                             from.Row, Math.Round(from.RowOff / EMU_PER_PIXEL * 1.5),
                             to.Column, Math.Round(to.ColumnOff / EMU_PER_PIXEL * 1.5),
                             to.Row, Math.Round(to.RowOff / EMU_PER_PIXEL * 1.5));
    }

    /// <summary>
    /// The object is printed when the document is printed.
    /// </summary>
    public override bool Print
    {
        get
        {
            return this._control.Print;
        }
        set
        {
            this._control.Print = value;
            base.Print = value;
        }
    }

    /// <summary>
    /// The object is locked when the sheet is protected..
    /// </summary>
    public override bool Locked
    {
        get
        {
            return this._control.Locked;
        }
        set
        {
            this._control.Locked = value;
            base.Locked = value;
        }
    }
    /// <summary>
    /// If the controls fill formatting is provided automatically
    /// </summary>
    public bool AutoFill
    {
        get { return this._control.AutoFill; }
        set { this._control.AutoFill = value; }
    }

    /// <summary>
    /// If the controls size is formatted automatically.
    /// </summary>
    public bool AutoPict
    {
        get { return this._control.AutoPict; }
        set { this._control.AutoPict = value; }
    }

    /// <summary>
    /// Returns true if the object is at its default size.
    /// </summary>
    public bool DefaultSize
    {
        get { return this._control.DefaultSize; }
        set { this._control.DefaultSize = value; }
    }

    /// <summary>
    /// If true, the object is allowed to run an attached macro
    /// </summary>
    public bool Disabled
    {
        get { return this._control.Disabled; }
        set { this._control.Disabled = value; }
    }
    /// <summary>
    /// If the control has 3D effects enabled.
    /// </summary>
    public bool ThreeDEffects
    {
        get
        {
            bool? b = this._ctrlProp.GetXmlNodeBoolNullable("@noThreeD2");
            if (b.HasValue == false)
            {
                return this._ctrlProp.GetXmlNodeBool("@noThreeD") == false;
            }
            else
            {
                return this._ctrlProp.GetXmlNodeBool("@noThreeD2") == false;
            }
        }
        set
        {
            bool? b = this._ctrlProp.GetXmlNodeBoolNullable("@noThreeD2");
            if (b.HasValue)
            {
                this._ctrlProp.SetXmlNodeBool("@noThreeD2", value == false);   //can be used for lists and drop-downs.
            }
            else
            {
                this._ctrlProp.SetXmlNodeBool("@noThreeD", value == false);
            }

            string? xmlAttr = (this.ControlType == eControlType.DropDown || this.ControlType == eControlType.ListBox) ? "x:NoThreeD2" : "x:NoThreeD";
            if (value)
            {
                this._vmlProp.CreateNode(xmlAttr);
            }
            else
            {
                this._vmlProp.DeleteNode(xmlAttr);
            }
        }
    }
    /// <summary>
    /// Gets or sets the address to the cell that is linked to the control. 
    /// </summary>
    public ExcelAddressBase LinkedCell
    {
        get
        {
            if (this.ControlType == eControlType.Label || this.ControlType == eControlType.Button || this.ControlType == eControlType.GroupBox)
            {
                return this.FmlaTxbx;
            }
            else
            {
                return this.FmlaLink;
            }
        }
        set
        {
            if (this.ControlType == eControlType.Label || this.ControlType == eControlType.Button || this.ControlType == eControlType.GroupBox)
            {
                this.FmlaTxbx = value;
            }
            else
            {
                this.FmlaLink = value;
            }
        }
    }
    internal void SetLinkedCellValue(int value)
    {
        if (this.LinkedCell != null)
        {
            ExcelWorksheet ws;
            if (string.IsNullOrEmpty(this.LinkedCell.WorkSheetName))
            {
                ws = this._drawings.Worksheet;
            }
            else
            {
                ws = this._drawings.Worksheet.Workbook.Worksheets[this.LinkedCell.WorkSheetName];
            }
            ws.Cells[this.LinkedCell._fromRow, this.LinkedCell._fromCol].Value = value;
        }
    }
        
    #region Shared Properties
    internal ExcelAddressBase FmlaLink
    {
        get
        {
            string? range = this._ctrlProp.GetXmlNodeString("@fmlaLink");
            if (ExcelCellBase.IsValidAddress(range))
            {
                return new ExcelAddressBase(range);
            }
            return null;
        }
        set
        {
            if (value == null)
            {
                this._ctrlProp.DeleteNode("@fmlaLink");
                this._vmlProp.DeleteNode("x:FmlaLink");
            }
            if (value.WorkSheetName.Equals(this._drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
            {
                this._ctrlProp.SetXmlNodeString("@fmlaLink", value.Address);
                this._vmlProp.SetXmlNodeString("x:FmlaLink", value.Address);
            }
            else
            {
                this._ctrlProp.SetXmlNodeString("@fmlaLink", value.FullAddress);
                this._vmlProp.SetXmlNodeString("x:FmlaLink", value.FullAddress);
            }
        }
    }
    /// <summary>
    /// The source data cell that the control object's data is linked to.
    /// </summary>
    internal ExcelAddressBase FmlaTxbx
    {
        get
        {
            string? range = this._ctrlProp.GetXmlNodeString("@fmlaTxbx");
            if (ExcelCellBase.IsValidAddress(range))
            {
                return new ExcelAddressBase(range);
            }
            return null;
        }
        set
        {
            if (value == null)
            {
                this._ctrlProp.DeleteNode("@fmlaTxbx");
                this._vmlProp.DeleteNode("x:FmlaTxbx");
            }
            if (value.WorkSheetName.Equals(this._drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
            {
                this._ctrlProp.SetXmlNodeString("@fmlaTxbx", value.Address);
                this._vmlProp.SetXmlNodeString("x:FmlaTxbx", value.Address);
            }
            else
            {
                this._ctrlProp.SetXmlNodeString("@fmlaTxbx", value.FullAddress);
                this._vmlProp.SetXmlNodeString("x:FmlaTxbx", value.FullAddress);
            }
        }
    }
    internal ExcelAddressBase LinkedGroup
    {
        get
        {
            string? range = this._ctrlProp.GetXmlNodeString("@fmlaGroup");
            if (ExcelCellBase.IsValidAddress(range))
            {
                return new ExcelAddressBase(range);
            }
            return null;
        }
        set
        {
            if (value == null)
            {
                this._ctrlProp.DeleteNode("@fmlaGroup");
                this._vmlProp.DeleteNode("x:FmlaGroup");
            }
            if (value.WorkSheetName.Equals(this._drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
            {
                this._ctrlProp.SetXmlNodeString("@fmlaGroup", value.Address);
                this._vmlProp.SetXmlNodeString("x:FmlaGroup", value.Address);
            }
            else
            {
                this._ctrlProp.SetXmlNodeString("@fmlaGroup", value.FullAddress);
                this._vmlProp.SetXmlNodeString("x:FmlaGroup", value.FullAddress);
            }
        }
    }
    /// <summary>
    /// The type of drawing. Always set to <see cref="eDrawingType.Control"/>
    /// </summary>
    public override eDrawingType DrawingType
    {
        get
        {
            return eDrawingType.Control;
        }
    }

    internal virtual void UpdateXml()
    {
        this.SetPositionAndSizeForControl();
        if(this.ControlType==eControlType.CheckBox || this.ControlType == eControlType.RadioButton)
        {
            ExcelControlWithColorsAndLines? c = (ExcelControlWithColorsAndLines)this;
                
            if(c.Fill.Style!=eVmlFillType.NoFill)
            {
                ExcelDrawingFill? fill = new ExcelDrawingFill(this._drawings, this.NameSpaceManager, this.TopNode, this._topPath+"/xdr:spPr", this.SchemaNodeOrder);
                if(c.Fill.Style==eVmlFillType.Solid) //Set solid fill for drawing. 
                {
                    Color color = c.Fill.Color.GetColor();
                    if (!color.IsEmpty)
                    {
                        fill.Color = color;
                    }
                    fill.Transparancy = (int)c.Fill.Opacity - 100;
                }
            }
        }
    }

    private void SetPositionAndSizeForControl()
    {
        if (this.Position == null)
        {
            this._control.From.Row = this.From.Row;
            this._control.From.RowOff = this.From.RowOff;
            this._control.From.Column = this.From.Column;
            this._control.From.ColumnOff = this.From.ColumnOff;
        }
        else
        {
            this.CalcColFromPixelLeft(this._left, out int col, out int colOff);
            this._control.From.Column = col;
            this._control.From.ColumnOff = colOff;

            this.CalcRowFromPixelTop(this._top, out int row, out int rowOff);
            this._control.From.Row = row;
            this._control.From.RowOff = rowOff;
        }

        if (this.Size == null)
        {
            this._control.To.Row = this.To.Row;
            this._control.To.RowOff = this.To.RowOff;
            this._control.To.Column = this.To.Column;
            this._control.To.ColumnOff = this.To.ColumnOff;
        }
        else
        {
            this.GetToRowFromPixels(this._height, out int row, out int rowOff, this._control.From.Row, this._control.From.RowOff);
            this.GetToColumnFromPixels(this._width, out int col, out int pixOff, this._control.From.Column, this._control.From.ColumnOff);
            this._control.To.Row = row;
            this._control.To.RowOff = rowOff;

            this._control.To.Column = col - 2;
            this._control.To.ColumnOff = pixOff * EMU_PER_PIXEL;
        }

        if (this._parent == null)
        {
            this._control.MoveWithCells = this.EditAs != eEditAs.Absolute;
            this._control.SizeWithCells = this.EditAs == eEditAs.TwoCell;
        }

        this._control.From.UpdateXml();
        this._control.To.UpdateXml();

        this._vml.Anchor = this.GetVmlAnchorValue();
    }

    internal static eEditAs GetControlEditAs(eControlType controlType)
    {
        switch(controlType)
        {
            case eControlType.Button:
                return eEditAs.Absolute;
            case eControlType.Label:
                return eEditAs.TwoCell;
            default:
                return eEditAs.OneCell;
        }
    }

    #endregion
    internal override void DeleteMe()
    {
        this._vml.TopNode.ParentNode.RemoveChild(this._vml.TopNode);
        this._drawings._package.ZipPackage.DeletePart(this.ControlPropertiesUri);
        this._control.DeleteMe();
        base.DeleteMe();
    }
}