/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/01/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /*
    <xsd:complexType name="CT_Control">
    3135 <xsd:sequence>
    3136 <xsd:element name="controlPr" type="CT_ControlPr" minOccurs="0" maxOccurs="1"/>
    3137 </xsd:sequence>
    3138 <xsd:attribute name="shapeId" type="xsd:unsignedInt" use="required"/>
    3139 <xsd:attribute ref="r:id" use="required"/>
    3140 <xsd:attribute name="name" type="xsd:string" use="optional"/>
    3141 </xsd:complexType>
    3142 <xsd:complexType name="CT_ControlPr">
    3143 <xsd:sequence>
    3144 <xsd:element name="anchor" type="CT_ObjectAnchor" minOccurs="1" maxOccurs="1"/>
    3145 </xsd:sequence>
    3146 <xsd:attribute name="locked" type="xsd:boolean" use="optional" default="true"/>
    3147 <xsd:attribute name="defaultSize" type="xsd:boolean" use="optional" default="true"/>
    3148 <xsd:attribute name="print" type="xsd:boolean" use="optional" default="true"/>
    3149 <xsd:attribute name="disabled" type="xsd:boolean" use="optional" default="false"/>
    3150 <xsd:attribute name="recalcAlways" type="xsd:boolean" use="optional" default="false"/>
    3151 <xsd:attribute name="uiObject" type="xsd:boolean" use="optional" default="false"/>
    3152 <xsd:attribute name="autoFill" type="xsd:boolean" use="optional" default="true"/>
    3153 <xsd:attribute name="autoLine" type="xsd:boolean" use="optional" default="true"/>
    3154 <xsd:attribute name="autoPict" type="xsd:boolean" use="optional" default="true"/>
    3155 <xsd:attribute name="macro" type="ST_Formula" use="optional"/>
    3156 <xsd:attribute name="altText" type="s:ST_Xstring" use="optional"/>
    3157 <xsd:attribute name="linkedCell" type="ST_Formula" use="optional"/>
    3158 <xsd:attribute name="listFillRange" type="ST_Formula" use="optional"/>
    3159 <xsd:attribute name="cf" type="s:ST_Xstring" use="optional" default="pict"/>
    3160 <xsd:attribute ref="r:id" use="optional"/>
    3161 </xsd:complexType> 
    */
    internal class ControlInternal : XmlHelper
    {

        internal ControlInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {

        }

        public string RelationshipId 
        {
            get
            {
                return this.GetXmlNodeString("@r:id");
            }
            set
            {
                this.SetXmlNodeString("@r:id", value);
            }
        }

        public string Macro 
        { 
            get
            {
                return this.GetXmlNodeString("d:controlPr/@macro");
            }
            internal set
            {
                this.SetXmlNodeString("d:controlPr/@macro", value);                
            }
        }

        public bool Print
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@print", true);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@print", value);
            }
        }

        public bool Locked
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@locked", true);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@locked", value);
            }
        }



        public bool AutoPict
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@autoPict", true);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@autoPict", value);
            }
        }

        public bool AutoFill
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@autoFill", true);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@autoFill", value);
            }
        }

        public bool DefaultSize
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@defaultSize", true);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@defaultSize", value);
            }
        }

        public bool Disabled
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/@disabled", false);
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/@disabled", value);
            }
        }

        internal string Name 
        { 
            get
            {
                return this.GetXmlNodeString("@name");
            }
            set
            {
                this.SetXmlNodeString("@name", value);
            }
        }
        internal int ShapeId
        {
            get
            {
                return this.GetXmlNodeInt("@shapeId");
            }
            set
            {
                this.SetXmlNodeInt("@shapeId", value);
            }
        }
        internal string AlternativeText
        {
            get
            {
                return ConvertUtil.ExcelDecodeString(this.GetXmlNodeString("d:controlPr/@altText"));
            }
            set
            {
                this.SetXmlNodeString("d:controlPr/@altText", ConvertUtil.ExcelEncodeString(value));
            }
        }
        public string FormulaRange
        {
            get
            {
                return this.GetXmlNodeString("d:controlPr/@fmlaRange");
            }
            set
            {
                this.SetXmlNodeString("d:controlPr/@fmlaRange", value);
            }
        }
        public string LinkedCell
        {
            get
            {
                return this.GetXmlNodeString("d:controlPr/@linkedCell");
            }
            set
            {
                this.SetXmlNodeString("d:controlPr/@linkedCell", value);
            }
        }
        ExcelPosition _from = null;
        public ExcelPosition From
        {
            get { return this._from ?? (this._from = new ExcelPosition(this.NameSpaceManager, this.GetNode("d:controlPr/d:anchor/d:from"), null)); }
        }
        ExcelPosition _to=null;
        public ExcelPosition To
        {
            get { return this._to ?? (this._to = new ExcelPosition(this.NameSpaceManager, this.GetNode("d:controlPr/d:anchor/d:to"), null)); }
        }
        public bool MoveWithCells 
        { 
            get
            {
                return this.GetXmlNodeBool("d:controlPr/d:anchor/@moveWithCells");
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/d:anchor/@moveWithCells", value, false);
            }
        }
        public bool SizeWithCells
        {
            get
            {
                return this.GetXmlNodeBool("d:controlPr/d:anchor/@sizeWithCells");
            }
            set
            {
                this.SetXmlNodeBool("d:controlPr/d:anchor/@sizeWithCells", value, false);
            }
        }

        internal void DeleteMe()
        {
            XmlNode? node = this.TopNode.ParentNode?.ParentNode;
            if (node?.LocalName=="AlternateContent")
            {
                node.ParentNode.RemoveChild(node);
            }
            else
            {
                this.TopNode.ParentNode.RemoveChild(this.TopNode);
            }
        }
    }
}