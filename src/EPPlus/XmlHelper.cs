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
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml;

/// <summary>
/// Abstract helper class containing functionality to work with XML inside the package. 
/// </summary>
public abstract class XmlHelper
{
    int[] _levels;

    internal delegate int ChangedEventHandler(StyleBase sender, StyleChangeEventArgs e);

    internal XmlHelper(XmlNamespaceManager nameSpaceManager)
    {
        this.TopNode = null;
        this.NameSpaceManager = nameSpaceManager;
    }

    internal XmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
    {
        this.TopNode = topNode;
        this.NameSpaceManager = nameSpaceManager;
    }

    //internal bool ChangedFlag;
    internal XmlNamespaceManager NameSpaceManager { get; set; }

    internal XmlNode TopNode { get; set; }

    /// <summary>
    /// Schema order list
    /// </summary>
    internal string[] SchemaNodeOrder { get; set; }

    /// <summary>
    /// Adds a new array to the end of SchemaNodeOrder
    /// </summary>
    /// <param name="schemaNodeOrder">The order to start from </param>
    /// <param name="newItems">The new items</param>
    /// <returns>The new order</returns>
    internal void AddSchemaNodeOrder(string[] schemaNodeOrder, string[] newItems) => this.SchemaNodeOrder = CopyToSchemaNodeOrder(schemaNodeOrder, newItems);

    internal void SetBoolNode(string path, bool value)
    {
        if (value)
        {
            _ = this.CreateNode(path);
        }
        else
        {
            this.DeleteNode(path);
        }
    }

    /// <summary>
    /// Adds a new array to the end of SchemaNodeOrder
    /// </summary>
    /// <param name="schemaNodeOrder">The order to start from </param>
    /// <param name="newItems">The new items</param>
    /// <param name="levels">Positions that defines levels in the xpath</param>
    internal void AddSchemaNodeOrder(string[] schemaNodeOrder, string[] newItems, int[] levels)
    {
        this._levels = levels;
        this.SchemaNodeOrder = CopyToSchemaNodeOrder(schemaNodeOrder, newItems);
    }

    internal static string[] CopyToSchemaNodeOrder(string[] schemaNodeOrder, string[] newItems)
    {
        if (schemaNodeOrder == null)
        {
            return newItems;
        }
        else
        {
            string[]? newOrder = new string[schemaNodeOrder.Length + newItems.Length];
            Array.Copy(schemaNodeOrder, newOrder, schemaNodeOrder.Length);
            Array.Copy(newItems, 0, newOrder, schemaNodeOrder.Length, newItems.Length);

            return newOrder;
        }
    }

    internal static void CopyElement(XmlElement fromElement, XmlElement toElement, string[] ignoreAttribute = null)
    {
        toElement.InnerXml = fromElement.InnerXml;

        //if (ignoreAttribute == null) return;
        foreach (XmlAttribute a in fromElement.Attributes)
        {
            if (ignoreAttribute == null || !ignoreAttribute.Contains(a.LocalName))
            {
                if (string.IsNullOrEmpty(a.NamespaceURI))
                {
                    toElement.SetAttribute(a.Name, a.Value);
                }
                else
                {
                    _ = toElement.SetAttribute(a.LocalName, a.NamespaceURI, a.Value);
                }
            }
        }
    }

    internal XmlNode CreateNode(string path)
    {
        if (path == "")
        {
            return this.TopNode;
        }
        else
        {
            return this.CreateNode(path, false);
        }
    }

    internal XmlNode CreateNode(XmlNode node, string path)
    {
        if (path == "")
        {
            return node;
        }
        else
        {
            return this.CreateNode(node, path, false, false, "");
        }
    }

    internal XmlNode CreateNode(XmlNode node, string path, bool addNew)
    {
        if (path == "")
        {
            return node;
        }
        else
        {
            return this.CreateNode(node, path, false, addNew, "");
        }
    }

    /// <summary>
    /// Create the node path. Nodes are inserted according to the Schema node order
    /// </summary>
    /// <param name="path">The path to be created</param>
    /// <param name="insertFirst">Insert as first child</param>
    /// <param name="addNew">Always add a new item at the last level.</param>
    /// <param name="exitName">Exit if after this named node has been created</param>
    /// <returns></returns>
    internal XmlNode CreateNode(string path, bool insertFirst, bool addNew = false, string exitName = "") => this.CreateNode(this.TopNode, path, insertFirst, addNew, exitName);

    internal XmlNode CreateAlternateContentNode(string elementName, string requires) => this.CreateNode(this.TopNode, elementName, false, false, "", requires);

    private XmlNode CreateNode(XmlNode node, string path, bool insertFirst, bool addNew, string exitName, string alternateContentRequires = null)
    {
        XmlNode prependNode = null;
        int lastUsedOrderIndex = 0;

        if (path.StartsWith("/", StringComparison.OrdinalIgnoreCase))
        {
            path = path.Substring(1);
        }

        string[]? subPaths = path.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

        for (int i = 0; i < subPaths.Length; i++)
        {
            string subPath = subPaths[i];
            XmlNode subNode = node.SelectSingleNode(subPath, this.NameSpaceManager);

            if (subNode == null || (i == subPaths.Length - 1 && addNew))
            {
                string nodeName;
                string nodePrefix;

                string[] nameSplit = subPath.Split(':');

                if (this.SchemaNodeOrder != null && subPath[0] != '@')
                {
                    insertFirst = false;
                    prependNode = this.GetPrependNode(subPath, node, ref lastUsedOrderIndex);
                }


                string nameSpaceURI;
                if (nameSplit.Length > 1)
                {
                    nodePrefix = nameSplit[0];

                    if (nodePrefix[0] == '@')
                    {
                        nodePrefix = nodePrefix.Substring(1, nodePrefix.Length - 1);
                    }

                    nameSpaceURI = this.NameSpaceManager.LookupNamespace(nodePrefix);
                    nodeName = nameSplit[1];
                }
                else
                {
                    nodePrefix = "";
                    nameSpaceURI = "";
                    nodeName = nameSplit[0];
                }

                if (subPath.StartsWith("@", StringComparison.OrdinalIgnoreCase))
                {
                    XmlAttribute addedAtt = node.OwnerDocument.CreateAttribute(subPath.Substring(1, subPath.Length - 1), nameSpaceURI); //nameSpaceURI
                    _ = node.Attributes.Append(addedAtt);
                }
                else
                {
                    if (nodePrefix == "")
                    {
                        subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                    }
                    else
                    {
                        if (nodePrefix == ""
                            || (node.OwnerDocument != null
                                && node.OwnerDocument.DocumentElement != null
                                && node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI
                                && node.OwnerDocument.DocumentElement.Prefix == ""))
                        {
                            subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                        }
                        else
                        {
                            subNode = node.OwnerDocument.CreateElement(nodePrefix, nodeName, nameSpaceURI);
                        }
                    }

                    if (string.IsNullOrEmpty(alternateContentRequires) == false)
                    {
                        XmlElement? altNode = node.OwnerDocument.CreateElement("AlternateContent", ExcelPackage.schemaMarkupCompatibility);
                        XmlElement? choiceNode = node.OwnerDocument.CreateElement("Choice", ExcelPackage.schemaMarkupCompatibility);
                        _ = altNode.AppendChild(choiceNode);
                        choiceNode.SetAttribute("Requires", alternateContentRequires);
                        _ = choiceNode.AppendChild(subNode);
                        subNode = altNode;
                    }

                    if (prependNode != null)
                    {
                        _ = node.InsertBefore(subNode, prependNode);
                        prependNode = null;
                    }
                    else if (insertFirst || (this.SchemaNodeOrder?.Length > 0 && subNode.LocalName == this.SchemaNodeOrder[0]))
                    {
                        _ = node.PrependChild(subNode);
                    }
                    else
                    {
                        _ = node.AppendChild(subNode);
                    }
                }

                if (nodeName == exitName)
                {
                    return subNode;
                }
            }
            else if
                (this.SchemaNodeOrder != null
                 && subPath != "..") //Parent node, node order should not change. Parent node (..) is only supported in the start of the xpath
            {
                int ix = this.GetNodePos(subNode.LocalName, lastUsedOrderIndex);

                if (ix >= 0)
                {
                    lastUsedOrderIndex = this.GetIndex(ix);
                }
            }

            node = subNode;
        }

        return node;
    }

    internal bool CreateNodeUntil(string path, string untilNodeName, out XmlNode spPrNode)
    {
        spPrNode = this.CreateNode(path, false, false, untilNodeName);

        return spPrNode != null && spPrNode.LocalName == untilNodeName;
    }

    internal XmlNode ReplaceElement(XmlNode oldChild, string newNodeName)
    {
        string[]? newNameSplit = newNodeName.Split(':');
        XmlElement newElement;

        if (newNodeName.Length > 1)
        {
            string? prefix = newNameSplit[0];

            string? ns = this.NameSpaceManager.LookupNamespace(prefix);
            newElement = oldChild.OwnerDocument.CreateElement(newNodeName, ns);
        }
        else
        {
            newElement = oldChild.OwnerDocument.CreateElement(newNodeName, this.NameSpaceManager.DefaultNamespace);
        }

        _ = oldChild.ParentNode.ReplaceChild(newElement, oldChild);

        return newElement;
    }

    /// <summary>
    /// Options to insert a node in the XmlDocument
    /// </summary>
    internal enum eNodeInsertOrder
    {
        /// <summary>
        /// Insert as first node of "topNode"
        /// </summary>
        First,

        /// <summary>
        /// Insert as the last child of "topNode"
        /// </summary>
        Last,

        /// <summary>
        /// Insert after the "referenceNode"
        /// </summary>
        After,

        /// <summary>
        /// Insert before the "referenceNode"
        /// </summary>
        Before,

        /// <summary>
        /// Use the Schema List to insert in the right order. If the Schema list
        /// is null or empty, consider "Last" as the selected option
        /// </summary>
        SchemaOrder
    }

    /// <summary>
    /// Create a complex node. Insert the node according to SchemaOrder
    /// using the TopNode as the parent
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    internal XmlNode CreateComplexNode(string path) => this.CreateComplexNode(this.TopNode, path, eNodeInsertOrder.SchemaOrder, null);

    /// <summary>
    /// Create a complex node. Insert the node according to the <paramref name="path"/>
    /// using the <paramref name="topNode"/> as the parent
    /// </summary>
    /// <param name="topNode"></param>
    /// <param name="path"></param>
    /// <returns></returns>
    internal XmlNode CreateComplexNode(XmlNode topNode, string path) => this.CreateComplexNode(topNode, path, eNodeInsertOrder.SchemaOrder, null);

    /// <summary>
    /// Creates complex XML nodes
    /// </summary>
    /// <remarks>
    /// 1. "d:conditionalFormatting"
    ///     1.1. Creates/find the first "conditionalFormatting" node
    /// 
    /// 2. "d:conditionalFormatting/@sqref"
    ///     2.1. Creates/find the first "conditionalFormatting" node
    ///     2.2. Creates (if not exists) the @sqref attribute
    ///
    /// 3. "d:conditionalFormatting/@id='7'/@sqref='A9:B99'"
    ///     3.1. Creates/find the first "conditionalFormatting" node
    ///     3.2. Creates/update its @id attribute to "7"
    ///     3.3. Creates/update its @sqref attribute to "A9:B99"
    ///
    /// 4. "d:conditionalFormatting[@id='7']/@sqref='X1:X5'"
    ///     4.1. Creates/find the first "conditionalFormatting" node with @id=7
    ///     4.2. Creates/update its @sqref attribute to "X1:X5"
    /// 
    /// 5. "d:conditionalFormatting[@id='7']/@id='8'/@sqref='X1:X5'/d:cfRule/@id='AB'"
    ///     5.1. Creates/find the first "conditionalFormatting" node with @id=7
    ///     5.2. Set its @id attribute to "8"
    ///     5.2. Creates/update its @sqref attribute and set it to "X1:X5"
    ///     5.3. Creates/find the first "cfRule" node (inside the node)
    ///     5.4. Creates/update its @id attribute to "AB"
    /// 
    /// 6. "d:cfRule/@id=''"
    ///     6.1. Creates/find the first "cfRule" node
    ///     6.1. Remove the @id attribute
    /// </remarks>
    /// <param name="topNode"></param>
    /// <param name="path"></param>
    /// <param name="nodeInsertOrder"></param>
    /// <param name="referenceNode"></param>
    /// <returns>The last node creates/found</returns>
    internal XmlNode CreateComplexNode(XmlNode topNode, string path, eNodeInsertOrder nodeInsertOrder, XmlNode referenceNode)
    {
        // Path is obrigatory
        if (path == null || path == string.Empty)
        {
            return topNode;
        }

        XmlNode node = topNode;
        int lastIndex = 0;

        //TODO: BUG: when the "path" contains "/" in an attrribue value, it gives an error.

        // Separate the XPath to Nodes and Attributes
        foreach (string subPath in path.Split('/'))
        {
            // The subPath can be any one of those:
            // nodeName
            // x:nodeName
            // nodeName[find criteria]
            // x:nodeName[find criteria]
            // @attribute
            // @attribute='attribute value'

            // Check if the subPath has at least one character
            if (subPath.Length > 0)
            {
                // Check if the subPath is an attribute (with or without value)
                if (subPath.StartsWith("@", StringComparison.OrdinalIgnoreCase))
                {
                    // @attribute                                       --> Create attribute
                    // @attribute=''                                --> Remove attribute
                    // @attribute='attribute value' --> Create attribute + update value
                    string[] attributeSplit = subPath.Split('=');
                    string attributeName = attributeSplit[0].Substring(1, attributeSplit[0].Length - 1);
                    string attributeValue = null; // Null means no attribute value

                    // Check if we have an attribute value to set
                    if (attributeSplit.Length > 1)
                    {
                        // Remove the ' or " from the attribute value
                        attributeValue = attributeSplit[1].Replace("'", "").Replace("\"", "");
                    }

                    // Get the attribute (if exists)
                    XmlAttribute attribute = (XmlAttribute)node.Attributes.GetNamedItem(attributeName);

                    // Remove the attribute if value is empty (not null)
                    if (attributeValue == string.Empty)
                    {
                        // Only if the attribute exists
                        if (attribute != null)
                        {
                            _ = node.Attributes.Remove(attribute);
                        }
                    }
                    else
                    {
                        // Create the attribue if does not exists
                        if (attribute == null)
                        {
                            // Create the attribute
                            attribute = node.OwnerDocument.CreateAttribute(attributeName);

                            // Add it to the current node
                            _ = node.Attributes.Append(attribute);
                        }

                        // Update the attribute value
                        if (attributeValue != null)
                        {
                            node.Attributes[attributeName].Value = attributeValue;
                        }
                    }
                }
                else
                {
                    // nodeName
                    // x:nodeName
                    // nodeName[find criteria]
                    // x:nodeName[find criteria]

                    // Look for the node (with or without filter criteria)
                    XmlNode subNode = node.SelectSingleNode(subPath, this.NameSpaceManager);

                    // Check if the node does not exists
                    if (subNode == null)
                    {
                        string nodeName;
                        string nodePrefix;
                        string[] nameSplit = subPath.Split(':');
                        string nameSpaceURI;

                        // Check if the name has a prefix like "d:nodeName"
                        if (nameSplit.Length > 1)
                        {
                            nodePrefix = nameSplit[0];
                            nameSpaceURI = this.NameSpaceManager.LookupNamespace(nodePrefix);
                            nodeName = nameSplit[1];
                        }
                        else
                        {
                            nodePrefix = string.Empty;
                            nameSpaceURI = string.Empty;
                            nodeName = nameSplit[0];
                        }

                        // Check if we have a criteria part in the node name
                        if (nodeName.IndexOf('[') > 0)
                        {
                            // remove the criteria from the node name
                            nodeName = nodeName.Substring(0, nodeName.IndexOf('['));
                        }

                        if (nodePrefix == string.Empty)
                        {
                            subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                        }
                        else
                        {
                            if (node.OwnerDocument != null
                                && node.OwnerDocument.DocumentElement != null
                                && node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI
                                && node.OwnerDocument.DocumentElement.Prefix == string.Empty)
                            {
                                subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                            }
                            else
                            {
                                subNode = node.OwnerDocument.CreateElement(nodePrefix, nodeName, nameSpaceURI);
                            }
                        }

                        // Check if we need to use the "SchemaOrder"
                        if (nodeInsertOrder == eNodeInsertOrder.SchemaOrder)
                        {
                            // Check if the Schema Order List is empty
                            if (this.SchemaNodeOrder == null || this.SchemaNodeOrder.Length == 0)
                            {
                                // Use the "Insert Last" option when Schema Order List is empty
                                nodeInsertOrder = eNodeInsertOrder.Last;
                            }
                            else
                            {
                                // Find the prepend node in order to insert
                                referenceNode = this.GetPrependNode(nodeName, node, ref lastIndex);

                                if (referenceNode != null)
                                {
                                    nodeInsertOrder = eNodeInsertOrder.Before;
                                }
                                else
                                {
                                    nodeInsertOrder = eNodeInsertOrder.Last;
                                }
                            }
                        }

                        switch (nodeInsertOrder)
                        {
                            case eNodeInsertOrder.After:
                                _ = node.InsertAfter(subNode, referenceNode);
                                referenceNode = null;

                                break;

                            case eNodeInsertOrder.Before:
                                _ = node.InsertBefore(subNode, referenceNode);
                                referenceNode = null;

                                break;

                            case eNodeInsertOrder.First:
                                _ = node.PrependChild(subNode);

                                break;

                            case eNodeInsertOrder.Last:
                                _ = node.AppendChild(subNode);

                                break;
                        }
                    }

                    // Make the newly created node the top node when the rest of the path
                    // is being evaluated. So newly created nodes will be the children of the
                    // one we just created.
                    node = subNode;
                }
            }
        }

        // Return the last created/found node
        return node;
    }

    internal XmlNode GetNode(string path) => this.TopNode.SelectSingleNode(path, this.NameSpaceManager);

    internal XmlNodeList GetNodes(string path) => this.TopNode.SelectNodes(path, this.NameSpaceManager);

    internal void ClearChildren(string path)
    {
        XmlNode? n = this.TopNode.SelectSingleNode(path, this.NameSpaceManager);

        if (n != null)
        {
            n.InnerXml = null;
        }
    }

    /// <summary>
    /// return Prepend node
    /// </summary>
    /// <param name="nodeName">name of the node to check</param>
    /// <param name="node">Topnode to check children</param>
    /// <param name="index">Out index to keep track of level in the xml</param>
    /// <returns></returns>
    private XmlNode GetPrependNode(string nodeName, XmlNode node, ref int index)
    {
        int ix = this.GetNodePos(nodeName, index);

        if (ix < 0)
        {
            return null;
        }

        XmlNode prependNode = null;

        foreach (XmlNode childNode in node.ChildNodes)
        {
            string checkNodeName;

            if (childNode.LocalName
                == "AlternateContent") //AlternateContent contains the node that should be in the correnct order. For example AlternateContent/Choice/controls
            {
                checkNodeName = childNode.FirstChild?.FirstChild?.Name;
            }
            else
            {
                checkNodeName = childNode.Name;
            }

            int childPos = this.GetNodePos(checkNodeName, index);

            if (childPos > -1) //Found?
            {
                if (childPos > ix) //Position is before
                {
                    index = childPos + 1;

                    return childNode;
                }
            }
        }

        index = this.GetIndex(ix + 1);

        return prependNode;
    }

    private int GetIndex(int ix)
    {
        if (this._levels != null)
        {
            for (int i = 0; i <= this._levels.GetUpperBound(0); i++)
            {
                if (this._levels[i] >= ix)
                {
                    return this._levels[i];
                }
            }
        }

        return ix;
    }

    private int GetNodePos(string nodeName, int startIndex)
    {
        int ix = nodeName.IndexOf(':');

        if (ix > 0)
        {
            nodeName = nodeName.Substring(ix + 1, nodeName.Length - (ix + 1));
        }

        for (int i = startIndex; i < this.SchemaNodeOrder.Length; i++)
        {
            if (nodeName == this.SchemaNodeOrder[i])
            {
                return i;
            }
        }

        return -1;
    }

    internal void DeleteAllNode(string path)
    {
        string[] split = path.Split('/');
        XmlNode node = this.TopNode;

        foreach (string s in split)
        {
            node = node.SelectSingleNode(s, this.NameSpaceManager);

            if (node != null)
            {
                if (node is XmlAttribute)
                {
                    _ = (node as XmlAttribute).OwnerElement.Attributes.Remove(node as XmlAttribute);
                }
                else
                {
                    _ = node.ParentNode.RemoveChild(node);
                }
            }
            else
            {
                break;
            }
        }
    }

    /// <summary>
    /// Delete the element or attribut matching the XPath
    /// </summary>
    /// <param name="path">The path</param>
    /// <param name="deleteElement">If true and the node is an attribute, the parent element is deleted. Default false</param>
    internal void DeleteNode(string path, bool deleteElement = false)
    {
        XmlNode? node = this.TopNode.SelectSingleNode(path, this.NameSpaceManager);

        if (node != null)
        {
            if (node is XmlAttribute)
            {
                XmlAttribute? att = (XmlAttribute)node;

                if (deleteElement)
                {
                    _ = att.OwnerElement.ParentNode.RemoveChild(att.OwnerElement);
                }
                else
                {
                    _ = att.OwnerElement.Attributes.Remove(att);
                }
            }
            else
            {
                _ = node.ParentNode.RemoveChild(node);
            }
        }
    }

    internal void DeleteTopNode() => _ = this.TopNode.ParentNode.RemoveChild(this.TopNode);

    internal void SetXmlNodeDouble(string path, double? d, bool allowNegative) => this.SetXmlNodeDouble(path, d, null, "", allowNegative);

    internal void SetXmlNodeDouble(string path, double? d, CultureInfo ci = null, string suffix = "", bool allowNegative = true)
    {
        if (d.HasValue == false || double.IsNaN(d.Value))
        {
            this.DeleteNode(path);
        }
        else
        {
            if (allowNegative == false && d.Value < 0)
            {
                throw new InvalidOperationException("Value can't be negative");
            }

            this.SetXmlNodeString(this.TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture) + suffix);
        }
    }

    internal void SetXmlNodeInt(string path, int? d, CultureInfo ci = null, bool allowNegative = true)
    {
        if (d == null)
        {
            this.DeleteNode(path);
        }
        else
        {
            if (allowNegative == false && d.Value < 0)
            {
                throw new ArgumentException("Negative value not permitted");
            }

            this.SetXmlNodeString(this.TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture));
        }
    }

    internal void SetXmlNodeLong(string path, long? d, CultureInfo ci = null, bool allowNegative = true)
    {
        if (d == null)
        {
            this.DeleteNode(path);
        }
        else
        {
            if (allowNegative == false && d.Value < 0)
            {
                throw new ArgumentException("Negative value not permitted");
            }

            this.SetXmlNodeString(this.TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture));
        }
    }

    readonly char[] _whiteSpaces = new char[] { '\t', '\n', '\r', ' ' };

    internal void SetXmlNodeStringPreserveWhiteSpace(string path, string value, bool removeIfBlank = false, bool insertFirst = false)
    {
        this.SetXmlNodeString(this.TopNode, path, value, removeIfBlank, insertFirst);

        if (value != null && value.Length > 0)
        {
            if (this._whiteSpaces.Contains(value[0]) || this._whiteSpaces.Contains(value[value.Length - 1]))
            {
                XmlNode? workNode = this.GetNode(path);

                if (workNode.NodeType == XmlNodeType.Attribute)
                {
                    workNode = workNode.ParentNode;
                }

                if (workNode.NodeType == XmlNodeType.Element)
                {
                    ((XmlElement)workNode).SetAttribute("xml:space", "preserve");
                }
            }
        }
    }

    internal void SetXmlNodeString(string path, string value) => this.SetXmlNodeString(this.TopNode, path, value, false, false);

    internal void SetXmlNodeString(string path, string value, bool removeIfBlank) => this.SetXmlNodeString(this.TopNode, path, value, removeIfBlank, false);

    internal void SetXmlNodeString(XmlNode node, string path, string value) => this.SetXmlNodeString(node, path, value, false, false);

    internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank) => this.SetXmlNodeString(node, path, value, removeIfBlank, false);

    internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
    {
        if (node == null)
        {
            return;
        }

        if (value == "" && removeIfBlank)
        {
            this.DeleteAllNode(path);
        }
        else
        {
            XmlNode nameNode = node.SelectSingleNode(path, this.NameSpaceManager);

            if (nameNode == null)
            {
                _ = this.CreateNode(path, insertFirst);
                nameNode = node.SelectSingleNode(path, this.NameSpaceManager);
            }

            //if (nameNode.InnerText != value) HasChanged();
            nameNode.InnerText = value;
        }
    }

    internal void SetXmlNodeBool(string path, bool value) => this.SetXmlNodeString(this.TopNode, path, value ? "1" : "0", false, false);

    internal void SetXmlNodeBoolVml(string path, bool value) => this.SetXmlNodeString(this.TopNode, path, value ? "t" : "f", false, false);

    internal void SetXmlNodeBool(string path, bool value, bool removeIf)
    {
        if (value == removeIf)
        {
            XmlNode? node = this.TopNode.SelectSingleNode(path, this.NameSpaceManager);

            if (node != null)
            {
                if (node is XmlAttribute attrib)
                {
                    XmlElement? elem = attrib.OwnerElement;
                    elem.RemoveAttribute(node.Name);
                }
                else
                {
                    _ = node.ParentNode.RemoveChild(node);
                }
            }
        }
        else
        {
            this.SetXmlNodeString(this.TopNode, path, value ? "1" : "0", false, false);
        }
    }

    internal void SetXmlNodePercentage(string path, double? value, bool allowNegative = true, double minMaxValue = 100D)
    {
        if (value.HasValue)
        {
            if (allowNegative == false && value < 0)
            {
                throw new ArgumentException("Negative percentage not allowed");
            }

            if (value < -minMaxValue || value > minMaxValue)
            {
                throw new ArgumentOutOfRangeException(nameof(value), $"Percentage out of range. Ranges from {(allowNegative ? 0 : -minMaxValue)}% to {minMaxValue}%");
            }

            this.SetXmlNodeString(path, ((int)(value.Value * 1000)).ToString(CultureInfo.InvariantCulture));
        }
        else
        {
            this.DeleteNode(path);
        }
    }

    internal void SetXmlNodeAngel(string path, double? value, string parameter = null, int minValue = 0, int maxValue = 360)
    {
        if (value.HasValue)
        {
            if (!string.IsNullOrEmpty(parameter) && (value < minValue || value > maxValue))
            {
                throw new ArgumentOutOfRangeException(parameter, $"Value must be between {minValue} and {maxValue}");
            }

            int v = (int)(value * 60000);
            this.SetXmlNodeString(path, v.ToString(CultureInfo.InvariantCulture));
        }
        else
        {
            this.DeleteNode(path);
        }
    }

    internal void SetXmlNodeEmuToPt(string path, double? value)
    {
        if (value.HasValue)
        {
            int v = (int)(value * Drawing.ExcelDrawing.EMU_PER_POINT);
            this.SetXmlNodeString(path, v.ToString());
        }
        else
        {
            this.DeleteNode(path);
        }
    }

    internal void SetXmlNodeFontSize(string path, double? value, string propertyName, bool AllowNegative = true)
    {
        if (value.HasValue)
        {
            if (AllowNegative)
            {
                if (value < 0 || value > 4000)
                {
                    throw new ArgumentOutOfRangeException(propertyName, "Fontsize must be between 0 and 4000");
                }
            }
            else
            {
                if (value < -4000 || value > 4000)
                {
                    throw new ArgumentOutOfRangeException(propertyName, "Fontsize must be between -4000 and 4000");
                }
            }

            this.SetXmlNodeString(path, ((double)value * 100).ToString(CultureInfo.InvariantCulture));
        }
        else
        {
            this.DeleteNode(path);
        }
    }

    internal bool ExistsNode(string path)
    {
        if (this.TopNode == null || this.TopNode.SelectSingleNode(path, this.NameSpaceManager) == null)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    internal bool ExistsNode(XmlNode node, string path)
    {
        if (node == null || node.SelectSingleNode(path, this.NameSpaceManager) == null)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    internal bool? GetXmlNodeBoolNullable(string path)
    {
        string? value = this.GetXmlNodeString(path);

        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        return this.GetXmlNodeBool(path);
    }

    internal bool? GetXmlNodeBoolNullableWithVal(string path)
    {
        XmlNode? node = this.GetNode(path);

        if (node == null)
        {
            return null;
        }

        XmlAttribute? value = node.Attributes["val"];

        if (value == null)
        {
            return true;
        }
        else
        {
            return value.Value == "1" || value.Value == "-1" || value.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase);
        }
    }

    internal bool GetXmlNodeBool(string path) => this.GetXmlNodeBool(path, false);

    internal bool GetXmlNodeBool(string path, bool blankValue)
    {
        string value = this.GetXmlNodeString(path);

        if (value == "1" || value == "-1" || value.StartsWith("t", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        else if (value == "")
        {
            return blankValue;
        }
        else
        {
            return false;
        }
    }

    internal static bool GetBoolFromString(string s) => s != null && (s == "1" || s == "-1" || s.Equals("true", StringComparison.OrdinalIgnoreCase));

    internal static bool GetBoolFromNullString(string s) => s != null && (s == "1" || s == "-1" || s.Equals("true", StringComparison.OrdinalIgnoreCase));

    internal int GetXmlNodeInt(string path, int defaultValue = int.MinValue)
    {
        if (int.TryParse(this.GetXmlNodeString(path), NumberStyles.Number, CultureInfo.InvariantCulture, out int i))
        {
            return i;
        }
        else
        {
            return defaultValue;
        }
    }

    internal double GetXmlNodeAngel(string path, double defaultValue = 0)
    {
        int a = this.GetXmlNodeInt(path);

        if (a < 0)
        {
            return defaultValue;
        }

        return a / 60000D;
    }

    internal double GetXmlNodeEmuToPt(string path)
    {
        long v = this.GetXmlNodeLong(path);

        if (v < 0)
        {
            return 0;
        }

        return (double)(v / (double)Drawing.ExcelDrawing.EMU_PER_POINT);
    }

    internal double? GetXmlNodeEmuToPtNull(string path)
    {
        long? v = this.GetXmlNodeLongNull(path);

        if (v == null)
        {
            return null;
        }

        return (double)(v / (double)Drawing.ExcelDrawing.EMU_PER_POINT);
    }

    internal int? GetXmlNodeIntNull(string path)
    {
        string s = this.GetXmlNodeString(path);

        if (s != "" && int.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out int i))
        {
            return i;
        }
        else
        {
            return null;
        }
    }

    internal long GetXmlNodeLong(string path)
    {
        string s = this.GetXmlNodeString(path);

        if (s != "" && long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out long l))
        {
            return l;
        }
        else
        {
            return long.MinValue;
        }
    }

    internal long? GetXmlNodeLongNull(string path)
    {
        string s = this.GetXmlNodeString(path);

        if (s != "" && long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out long l))
        {
            return l;
        }
        else
        {
            return null;
        }
    }

    internal decimal GetXmlNodeDecimal(string path)
    {
        if (decimal.TryParse(this.GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal d))
        {
            return d;
        }
        else
        {
            return 0;
        }
    }

    internal decimal? GetXmlNodeDecimalNull(string path)
    {
        if (decimal.TryParse(this.GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal d))
        {
            return d;
        }
        else
        {
            return null;
        }
    }

    internal double? GetXmlNodeDoubleNull(string path)
    {
        string s = this.GetXmlNodeString(path);

        if (s == "")
        {
            return null;
        }
        else
        {
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
            {
                return v;
            }
            else
            {
                return null;
            }
        }
    }

    internal double GetXmlNodeDouble(string path)
    {
        string s = this.GetXmlNodeString(path);

        if (s == "")
        {
            return double.NaN;
        }
        else
        {
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
            {
                return v;
            }
            else
            {
                return double.NaN;
            }
        }
    }

    internal string GetXmlNodeString(XmlNode node, string path)
    {
        if (node == null)
        {
            return "";
        }

        XmlNode nameNode = node.SelectSingleNode(path, this.NameSpaceManager);

        if (nameNode != null)
        {
            if (nameNode.NodeType == XmlNodeType.Attribute)
            {
                return nameNode.Value ?? "";
            }
            else
            {
                return nameNode.InnerText;
            }
        }
        else
        {
            return "";
        }
    }

    internal string GetXmlNodeString(string path) => this.GetXmlNodeString(this.TopNode, path);

    internal static Uri GetNewUri(Packaging.ZipPackage package, string sUri)
    {
        int id = 1;

        return GetNewUri(package, sUri, ref id);
    }

    internal static Uri GetNewUri(Packaging.ZipPackage package, string sUri, ref int id)
    {
        Uri uri = new Uri(string.Format(sUri, id), UriKind.Relative);

        while (package.PartExists(uri))
        {
            uri = new Uri(string.Format(sUri, ++id), UriKind.Relative);
        }

        return uri;
    }

    internal T? GetXmlEnumNull<T>(string path, T? defaultValue = null)
        where T : struct, Enum
    {
        string? v = this.GetXmlNodeString(path);

        if (string.IsNullOrEmpty(v))
        {
            return defaultValue;
        }
        else
        {
            return v.ToEnum(default(T));
        }
    }

    internal double? GetXmlNodePercentage(string path)
    {
        double d;
        string? p = this.GetXmlNodeString(path);

        if (p.EndsWith("%"))
        {
            if (double.TryParse(p.Substring(0, p.Length - 1), out d))
            {
                return d;
            }
            else
            {
                return null;
            }
        }
        else
        {
            if (double.TryParse(p, out d))
            {
                return d / 1000;
            }
            else
            {
                return null;
            }
        }
    }

    internal double GetXmlNodeFontSize(string path) => (this.GetXmlNodeDoubleNull(path) ?? 0) / 100;

    internal void RenameNode(XmlNode node, string prefix, string newName, string[] allowedChildren = null)
    {
        XmlDocument? doc = node.OwnerDocument;
        XmlElement? newNode = doc.CreateElement(prefix, newName, this.NameSpaceManager.LookupNamespace(prefix));

        while (this.TopNode.ChildNodes.Count > 0)
        {
            if (allowedChildren == null || allowedChildren.Contains(this.TopNode.ChildNodes[0].LocalName))
            {
                _ = newNode.AppendChild(this.TopNode.ChildNodes[0]);
            }
            else
            {
                _ = this.TopNode.RemoveChild(this.TopNode.ChildNodes[0]);
            }
        }

        _ = this.TopNode.ParentNode.ReplaceChild(newNode, this.TopNode);
        this.TopNode = newNode;
    }

    /// <summary>
    /// Insert the new node before any of the nodes in the comma separeted list
    /// </summary>
    /// <param name="parentNode">Parent node</param>
    /// <param name="beforeNodes">comma separated list containing nodes to insert after. Left to right order</param>
    /// <param name="newNode">The new node to be inserterd</param>
    internal static void InserAfter(XmlNode parentNode, string beforeNodes, XmlNode newNode)
    {
        string[] nodePaths = beforeNodes.Split(',');

        XmlNode insertAfter = null;

        foreach (XmlNode childNode in parentNode.ChildNodes)
        {
            if (nodePaths.Contains(childNode.Name))
            {
                insertAfter = childNode;
            }
        }

        if (insertAfter == null)
        {
            _ = parentNode.AppendChild(newNode);
        }
        else
        {
            _ = parentNode.InsertAfter(newNode, insertAfter);
        }
    }

    internal static void LoadXmlSafe(XmlDocument xmlDoc, Stream stream)
    {
        XmlReaderSettings settings = new XmlReaderSettings();

        //Disable entity parsing (to aviod xmlbombs, External Entity Attacks etc).
#if(NET35)
            settings.ProhibitDtd = true;
#else
        settings.DtdProcessing = DtdProcessing.Prohibit;
#endif
        XmlReader reader = XmlReader.Create(stream, settings);
        xmlDoc.Load(reader);
    }

    internal static void LoadXmlSafe(XmlDocument xmlDoc, string xml, Encoding encoding)
    {
        using MemoryStream? stream = RecyclableMemory.GetStream(encoding.GetBytes(xml));
        LoadXmlSafe(xmlDoc, stream);
    }

    internal void CreatespPrNode(string nodePath = "c:spPr", bool withLine = true)
    {
        if (!this.ExistsNode(nodePath))
        {
            XmlNode? node = this.CreateNode(nodePath);

            if (withLine)
            {
                node.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/>";
            }
            else
            {
                node.InnerXml = "<a:noFill/><a:effectLst/><a:sp3d/>";
            }
        }
    }

    internal XmlNode GetOrCreateExtLstSubNode(string uriGuid, string prefix, string[] uriOrder = null)
    {
        foreach (XmlElement node in this.GetNodes("d:extLst/d:ext"))
        {
            if (node.Attributes["uri"].Value.Equals(uriGuid, StringComparison.OrdinalIgnoreCase))
            {
                return node;
            }
        }

        XmlElement? extLst = (XmlElement)this.CreateNode("d:extLst");
        XmlElement prependChild = null;

        if (uriOrder != null)
        {
            foreach (object? child in extLst.ChildNodes)
            {
                if (child is XmlElement e)
                {
                    int uo1 = Array.IndexOf(uriOrder, e.GetAttribute("uri"));
                    int uo2 = Array.IndexOf(uriOrder, uriGuid);

                    if (uo1 > uo2)
                    {
                        prependChild = e;
                    }
                }
            }
        }

        XmlElement? newExt = this.TopNode.OwnerDocument.CreateElement("ext", ExcelPackage.schemaMain);

        if (!string.IsNullOrEmpty(prefix))
        {
            newExt.SetAttribute($"xmlns:{prefix}", this.NameSpaceManager.LookupNamespace(prefix));
        }

        newExt.SetAttribute("uri", uriGuid);

        if (prependChild == null)
        {
            _ = extLst.AppendChild(newExt);
        }
        else
        {
            _ = extLst.InsertBefore(newExt, prependChild);
        }

        return newExt;
    }
}