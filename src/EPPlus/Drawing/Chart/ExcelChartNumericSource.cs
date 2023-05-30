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

using System.Xml;
using System.Globalization;
using System;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A numeric source for a chart.
/// </summary>
public class ExcelChartNumericSource : XmlHelper
{
    string _path;
    XmlElement _sourceElement;
    //string _formatCodePath;

    internal ExcelChartNumericSource(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, string[] schemaNodeOrder)
        : base(nameSpaceManager, topNode)
    {
        this._path = path;
        //this._formatCodePath = $"{this._path}/c:numLit/c:formatCode";
        this.AddSchemaNodeOrder(schemaNodeOrder, new string[] { "formatCode", "ptCount", "pt" });
        this.SetSourceElement();

        if (this._sourceElement != null)
        {
            switch (this._sourceElement.LocalName)
            {
                case "numLit":
                    this._formatCode = this.GetXmlNodeString(this._path + "/c:numLit/c:formatCode");

                    break;

                case "numRef":
                    this._formatCode = this.GetXmlNodeString(this._path + "/c:numRef/c:numCache/c:formatCode");

                    break;
            }
        }
    }

    /// <summary>
    /// This can be an address, function or litterals.
    /// Litternals are formatted as a comma separated list surrounded by curly brackets, for example {1.0,2.0,3}. Please use a dot(.) as decimal sign.
    /// </summary>
    public string ValuesSource
    {
        get
        {
            if (this._sourceElement == null)
            {
                return "";
            }
            else if (this._sourceElement.LocalName == "numLit")
            {
                return this.GetNumLit();
            }
            else
            {
                return this.GetXmlNodeString($"{this._path}/c:numRef/c:f");
            }
        }
        set
        {
            if (this._sourceElement != null)
            {
                _ = this._sourceElement.ParentNode.RemoveChild(this._sourceElement);
            }

            value = value.Trim();

            if (value.StartsWith("=", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(1);
            }

            if (value.StartsWith("{", StringComparison.OrdinalIgnoreCase))
            {
                if (!value.EndsWith("}", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("ValueSource", "Invalid format:Litteral values must begin and end with a curly bracket");
                }

                this.CreateNumLit(value);
            }
            else
            {
                this.SetXmlNodeString($"{this._path}/c:numRef/c:f", value);
            }

            //if (!string.IsNullOrEmpty(this._formatCode))
            //{
            //    this.FormatCode = this.FormatCode;
            //}

            this.SetSourceElement();
        }
    }

    private string GetNumLit()
    {
        string? v = "";

        foreach (XmlNode node in this._sourceElement.ChildNodes)
        {
            if (node.LocalName == "pt")
            {
                v += node.FirstChild.InnerText + ",";
            }
        }

        if (v.Length > 0)
        {
            v = "{" + v.Substring(0, v.Length - 1) + "}";
        }

        return v;
    }

    private void SetSourceElement()
    {
        XmlNode? node = this.GetNode(this._path);

        if (node != null && node.HasChildNodes)
        {
            this._sourceElement = (XmlElement)node.FirstChild;
        }
    }

    private void CreateNumLit(string value)
    {
        string[]? nums = value.Substring(1, value.Length - 2).Split(',');

        if (nums.Length > 0)
        {
            this.SetXmlNodeString($"{this._path}/c:numLit/c:ptCount/@val", nums.Length.ToString(CultureInfo.InvariantCulture));
            XmlElement? litNode = (XmlElement)this.GetNode($"{this._path}/c:numLit");
            int idx = 0;

            foreach (string? num in nums)
            {
                XmlElement? child = this.CreateLit(num.Trim(), idx++);
                _ = litNode.AppendChild(child);
            }
        }
    }

    private XmlElement CreateLit(string num, int idx)
    {
        XmlElement ptNode = this.TopNode.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
        ptNode.SetAttribute("idx", idx.ToString(CultureInfo.InvariantCulture));
        XmlElement? vNode = this.TopNode.OwnerDocument.CreateElement("c", "v", ExcelPackage.schemaChart);
        vNode.InnerText = num;
        _ = ptNode.AppendChild(vNode);

        return ptNode;
    }

    string _formatCode = "";

    /// <summary>
    /// The format code for the numeric source
    /// </summary>
    public string FormatCode
    {
        get { return this._formatCode; }
        set
        {
            if (this._sourceElement != null)
            {
                switch (this._sourceElement.LocalName)
                {
                    case "numLit":
                        this.SetXmlNodeString(this._path + "/c:numLit/c:formatCode", value);

                        break;

                    case "numRef":
                        this.SetXmlNodeString(this._path + "/c:numRef/c:numCache/c:formatCode", value);

                        break;
                }
            }

            this._formatCode = value;
        }
    }
}