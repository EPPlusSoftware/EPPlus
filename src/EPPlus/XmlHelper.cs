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

namespace OfficeOpenXml
{
    /// <summary>
    /// Abstract helper class containing functionality to work with XML inside the package. 
    /// </summary>
    public abstract class XmlHelper
    {
        int[] _levels = null;
        internal delegate int ChangedEventHandler(StyleBase sender, Style.StyleChangeEventArgs e);

        internal XmlHelper(XmlNamespaceManager nameSpaceManager)
        {
            TopNode = null;
            NameSpaceManager = nameSpaceManager;
        }

        internal XmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        {
            TopNode = topNode;
            NameSpaceManager = nameSpaceManager;
        }
        //internal bool ChangedFlag;
        internal XmlNamespaceManager NameSpaceManager { get; set; }
        internal XmlNode TopNode { get; set; }

        /// <summary>
        /// Schema order list
        /// </summary>
        internal string[] SchemaNodeOrder { get; set; } = null;
        /// <summary>
        /// Adds a new array to the end of SchemaNodeOrder
        /// </summary>
        /// <param name="schemaNodeOrder">The order to start from </param>
        /// <param name="newItems">The new items</param>
        /// <returns>The new order</returns>
        internal void AddSchemaNodeOrder(string[] schemaNodeOrder, string[] newItems)
        {
            SchemaNodeOrder = CopyToSchemaNodeOrder(schemaNodeOrder, newItems);
        }

        internal void SetBoolNode(string path, bool value)
        {
            if(value)
            {
                CreateNode(path);
            }
            else
            {
                DeleteNode(path);
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
            _levels = levels;
            SchemaNodeOrder = CopyToSchemaNodeOrder(schemaNodeOrder, newItems);
        }

        internal static string[] CopyToSchemaNodeOrder(string[] schemaNodeOrder, string[] newItems)
        {
            if (schemaNodeOrder == null)
            {
                return newItems;
            }
            else
            {
                var newOrder = new string[schemaNodeOrder.Length + newItems.Length];
                Array.Copy(schemaNodeOrder, newOrder, schemaNodeOrder.Length);
                Array.Copy(newItems, 0, newOrder, schemaNodeOrder.Length, newItems.Length);
                return newOrder;
            }
        }
        internal static void CopyElement(XmlElement fromElement, XmlElement toElement, string[] ignoreAttribute=null)
        {
            toElement.InnerXml = fromElement.InnerXml;
            //if (ignoreAttribute == null) return;
            foreach (XmlAttribute a in fromElement.Attributes)
            {
                if (ignoreAttribute==null || !ignoreAttribute.Contains(a.LocalName))
                {
                    if(string.IsNullOrEmpty(a.NamespaceURI))
                    {
                        toElement.SetAttribute(a.Name, a.Value);
                    }
                    else
                    {
                        toElement.SetAttribute(a.LocalName, a.NamespaceURI, a.Value);
                    }
                }
            }
        }
        internal XmlNode CreateNode(string path)
        {
            if (path == "")
                return TopNode;
            else
                return CreateNode(path, false);
        }
        internal XmlNode CreateNode(XmlNode node, string path)
        {
            if (path == "")
                return node;
            else
                return CreateNode(node, path, false, false,"");
        }
        internal XmlNode CreateNode(XmlNode node, string path, bool addNew)
        {
            if (path == "")
                return node;
            else
                return CreateNode(node, path, false, addNew, "");
        }

        /// <summary>
        /// Create the node path. Nodes are inserted according to the Schema node order
        /// </summary>
        /// <param name="path">The path to be created</param>
        /// <param name="insertFirst">Insert as first child</param>
        /// <param name="addNew">Always add a new item at the last level.</param>
        /// <param name="exitName">Exit if after this named node has been created</param>
        /// <returns></returns>
        internal XmlNode CreateNode(string path, bool insertFirst, bool addNew = false, string exitName = "")
        {
            return CreateNode(TopNode, path, insertFirst, addNew, exitName);
        }
        internal XmlNode CreateAlternateContentNode(string elementName, string requires)
        {
            return CreateNode(TopNode, elementName, false, false,"", requires);
        }

        private XmlNode CreateNode(XmlNode node, string path, bool insertFirst, bool addNew, string exitName, string alternateContentRequires=null)
        {
            XmlNode prependNode = null;
            int lastUsedOrderIndex = 0;
            if (path.StartsWith("/", StringComparison.OrdinalIgnoreCase)) path = path.Substring(1);
            var subPaths = path.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < subPaths.Length; i++)
            {
                string subPath = subPaths[i];
                XmlNode subNode = node.SelectSingleNode(subPath, NameSpaceManager);
                if (subNode == null || (i == subPaths.Length - 1 && addNew))
                {
                    string nodeName;
                    string nodePrefix;

                    string nameSpaceURI = "";
                    string[] nameSplit = subPath.Split(':');

                    if (SchemaNodeOrder != null && subPath[0] != '@')
                    {
                        insertFirst = false;
                        prependNode = GetPrependNode(subPath, node, ref lastUsedOrderIndex);
                    }

                    if (nameSplit.Length > 1)
                    {
                        nodePrefix = nameSplit[0];
                        if (nodePrefix[0] == '@') nodePrefix = nodePrefix.Substring(1, nodePrefix.Length - 1);
                        nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
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
                        XmlAttribute addedAtt = node.OwnerDocument.CreateAttribute(subPath.Substring(1, subPath.Length - 1), nameSpaceURI);  //nameSpaceURI
                        node.Attributes.Append(addedAtt);
                    }
                    else
                    {
                        if (nodePrefix == "")
                        {
                            subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                        }
                        else
                        {
                            if (nodePrefix == "" || (node.OwnerDocument != null && node.OwnerDocument.DocumentElement != null && node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI &&
                                    node.OwnerDocument.DocumentElement.Prefix == ""))
                            {
                                subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                            }
                            else
                            {
                                subNode = node.OwnerDocument.CreateElement(nodePrefix, nodeName, nameSpaceURI);
                            }
                        }
                        if(string.IsNullOrEmpty(alternateContentRequires)==false)
                        {
                            var altNode = node.OwnerDocument.CreateElement("AlternateContent", ExcelPackage.schemaMarkupCompatibility);
                            var choiceNode = node.OwnerDocument.CreateElement("Choice", ExcelPackage.schemaMarkupCompatibility);
                            altNode.AppendChild(choiceNode);
                            choiceNode.SetAttribute("Requires", alternateContentRequires);
                            choiceNode.AppendChild(subNode);
                            subNode=altNode;
                        }

                        if (prependNode != null)
                        {
                            node.InsertBefore(subNode, prependNode);
                            prependNode = null;
                        }
                        else if (insertFirst)
                        {
                            node.PrependChild(subNode);
                        }
                        else
                        {
                            node.AppendChild(subNode);
                        }
                    }
                    if (nodeName == exitName)
                    {
                        return subNode;
                    }
                }
                else if (SchemaNodeOrder != null && subPath != "..")  //Parent node, node order should not change. Parent node (..) is only supported in the start of the xpath
                {
                    var ix = GetNodePos(subNode.LocalName, lastUsedOrderIndex);
                    if (ix >= 0)
                    {
                        lastUsedOrderIndex = GetIndex(ix);
                    }
                }
                node = subNode;
            }
            return node;
        }

        internal bool CreateNodeUntil(string path, string untilNodeName, out XmlNode spPrNode)
        {
            spPrNode = CreateNode(path, false, false, untilNodeName);
            return spPrNode != null && spPrNode.LocalName == untilNodeName;
        }
        internal XmlNode ReplaceElement(XmlNode oldChild, string newNodeName)
        {
            var newNameSplit = newNodeName.Split(':');
            XmlElement newElement;
            if (newNodeName.Length > 1)
            {
                var prefix = newNameSplit[0];
                var name = newNameSplit[1];

                var ns = NameSpaceManager.LookupNamespace(prefix);
                newElement = oldChild.OwnerDocument.CreateElement(newNodeName, ns);
            }
            else
            {
                newElement = oldChild.OwnerDocument.CreateElement(newNodeName, NameSpaceManager.DefaultNamespace);
            }
            oldChild.ParentNode.ReplaceChild(newElement, oldChild);
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
        internal XmlNode CreateComplexNode(
            string path)
        {
            return CreateComplexNode(
                TopNode,
                path,
                eNodeInsertOrder.SchemaOrder,
                null);
        }

        /// <summary>
        /// Create a complex node. Insert the node according to the <paramref name="path"/>
        /// using the <paramref name="topNode"/> as the parent
        /// </summary>
        /// <param name="topNode"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        internal XmlNode CreateComplexNode(
            XmlNode topNode,
            string path)
        {
            return CreateComplexNode(
                topNode,
                path,
                eNodeInsertOrder.SchemaOrder,
                null);
        }

        /// <summary>
        /// Creates complex XML nodes
        /// </summary>
        /// <remarks>
        /// 1. "d:conditionalFormatting"
        ///		1.1. Creates/find the first "conditionalFormatting" node
        /// 
        /// 2. "d:conditionalFormatting/@sqref"
        ///		2.1. Creates/find the first "conditionalFormatting" node
        ///		2.2. Creates (if not exists) the @sqref attribute
        ///
        /// 3. "d:conditionalFormatting/@id='7'/@sqref='A9:B99'"
        ///		3.1. Creates/find the first "conditionalFormatting" node
        ///		3.2. Creates/update its @id attribute to "7"
        ///		3.3. Creates/update its @sqref attribute to "A9:B99"
        ///
        /// 4. "d:conditionalFormatting[@id='7']/@sqref='X1:X5'"
        ///		4.1. Creates/find the first "conditionalFormatting" node with @id=7
        ///		4.2. Creates/update its @sqref attribute to "X1:X5"
        ///	
        /// 5. "d:conditionalFormatting[@id='7']/@id='8'/@sqref='X1:X5'/d:cfRule/@id='AB'"
        ///		5.1. Creates/find the first "conditionalFormatting" node with @id=7
        ///		5.2. Set its @id attribute to "8"
        ///		5.2. Creates/update its @sqref attribute and set it to "X1:X5"
        ///		5.3. Creates/find the first "cfRule" node (inside the node)
        ///		5.4. Creates/update its @id attribute to "AB"
        ///	
        /// 6. "d:cfRule/@id=''"
        ///		6.1. Creates/find the first "cfRule" node
        ///		6.1. Remove the @id attribute
        ///	</remarks>
        /// <param name="topNode"></param>
        /// <param name="path"></param>
        /// <param name="nodeInsertOrder"></param>
        /// <param name="referenceNode"></param>
        /// <returns>The last node creates/found</returns>
        internal XmlNode CreateComplexNode(
            XmlNode topNode,
            string path,
            eNodeInsertOrder nodeInsertOrder,
            XmlNode referenceNode)
        {
            // Path is obrigatory
            if ((path == null) || (path == string.Empty))
            {
                return topNode;
            }

            XmlNode node = topNode;
            string nameSpaceURI = string.Empty;
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
                        // @attribute										--> Create attribute
                        // @attribute=''								--> Remove attribute
                        // @attribute='attribute value' --> Create attribute + update value
                        string[] attributeSplit = subPath.Split('=');
                        string attributeName = attributeSplit[0].Substring(1, attributeSplit[0].Length - 1);
                        string attributeValue = null;   // Null means no attribute value

                        // Check if we have an attribute value to set
                        if (attributeSplit.Length > 1)
                        {
                            // Remove the ' or " from the attribute value
                            attributeValue = attributeSplit[1].Replace("'", "").Replace("\"", "");
                        }

                        // Get the attribute (if exists)
                        XmlAttribute attribute = (XmlAttribute)(node.Attributes.GetNamedItem(attributeName));

                        // Remove the attribute if value is empty (not null)
                        if (attributeValue == string.Empty)
                        {
                            // Only if the attribute exists
                            if (attribute != null)
                            {
                                node.Attributes.Remove(attribute);
                            }
                        }
                        else
                        {
                            // Create the attribue if does not exists
                            if (attribute == null)
                            {
                                // Create the attribute
                                attribute = node.OwnerDocument.CreateAttribute(
                                    attributeName);

                                // Add it to the current node
                                node.Attributes.Append(attribute);
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
                        XmlNode subNode = node.SelectSingleNode(subPath, NameSpaceManager);

                        // Check if the node does not exists
                        if (subNode == null)
                        {
                            string nodeName;
                            string nodePrefix;
                            string[] nameSplit = subPath.Split(':');
                            nameSpaceURI = string.Empty;

                            // Check if the name has a prefix like "d:nodeName"
                            if (nameSplit.Length > 1)
                            {
                                nodePrefix = nameSplit[0];
                                nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
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
                                    subNode = node.OwnerDocument.CreateElement(
                                        nodeName,
                                        nameSpaceURI);
                                }
                                else
                                {
                                    subNode = node.OwnerDocument.CreateElement(
                                        nodePrefix,
                                        nodeName,
                                        nameSpaceURI);
                                }
                            }

                            // Check if we need to use the "SchemaOrder"
                            if (nodeInsertOrder == eNodeInsertOrder.SchemaOrder)
                            {
                                // Check if the Schema Order List is empty
                                if ((SchemaNodeOrder == null) || (SchemaNodeOrder.Length == 0))
                                {
                                    // Use the "Insert Last" option when Schema Order List is empty
                                    nodeInsertOrder = eNodeInsertOrder.Last;
                                }
                                else
                                {
                                    // Find the prepend node in order to insert
                                    referenceNode = GetPrependNode(nodeName, node, ref lastIndex);

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
                                    node.InsertAfter(subNode, referenceNode);
                                    referenceNode = null;
                                    break;

                                case eNodeInsertOrder.Before:
                                    node.InsertBefore(subNode, referenceNode);
                                    referenceNode = null;
                                    break;

                                case eNodeInsertOrder.First:
                                    node.PrependChild(subNode);
                                    break;

                                case eNodeInsertOrder.Last:
                                    node.AppendChild(subNode);
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

        internal XmlNode GetNode(string path)
        {
            return TopNode.SelectSingleNode(path, NameSpaceManager);
        }
        internal XmlNodeList GetNodes(string path)
        {
            return TopNode.SelectNodes(path, NameSpaceManager);
        }
        internal void ClearChildren(string path)
        {
            var n=TopNode.SelectSingleNode(path, NameSpaceManager);
            if(n!=null)
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
            var ix = GetNodePos(nodeName, index);
            if (ix < 0)
            {
                return null;
            }
            XmlNode prependNode = null;
            foreach(XmlNode childNode in node.ChildNodes)
            {
                string checkNodeName;
                if (childNode.LocalName=="AlternateContent") //AlternateContent contains the node that should be in the correnct order. For example AlternateContent/Choice/controls
                {
                    checkNodeName = childNode.FirstChild?.FirstChild?.Name;
                }
                else
                {
                    checkNodeName = childNode.Name;
                }
                int childPos = GetNodePos(checkNodeName, index);
                if (childPos > -1)  //Found?
                {
                    if (childPos > ix) //Position is before
                    {

                        index = childPos + 1;
                        return childNode;
                    }
                }
            }
            index = GetIndex(ix + 1);
            return prependNode;
        }

        private int GetIndex(int ix)
        {
            if (_levels != null)
            {
                for (int i = 0; i <= _levels.GetUpperBound(0); i++)
                {
                    if (_levels[i] >= ix)
                    {
                        return _levels[i];
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
            for (int i = startIndex; i < SchemaNodeOrder.Length; i++)
            {
                if (nodeName == SchemaNodeOrder[i])
                {
                    return i;
                }
            }
            return -1;
        }
        internal void DeleteAllNode(string path)
        {
            string[] split = path.Split('/');
            XmlNode node = TopNode;
            foreach (string s in split)
            {
                node = node.SelectSingleNode(s, NameSpaceManager);
                if (node != null)
                {
                    if (node is XmlAttribute)
                    {
                        (node as XmlAttribute).OwnerElement.Attributes.Remove(node as XmlAttribute);
                    }
                    else
                    {
                        node.ParentNode.RemoveChild(node);
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
		internal void DeleteNode(string path, bool deleteElement=false)
        {
            var node = TopNode.SelectSingleNode(path, NameSpaceManager);
            if (node != null)
            {
                if (node is XmlAttribute)
				{
                    var att = (XmlAttribute)node;
					if (deleteElement)
					{
						att.OwnerElement.ParentNode.RemoveChild(att.OwnerElement);
					}
					else
					{
						att.OwnerElement.Attributes.Remove(att);
					}
                }
                else
                {
                    node.ParentNode.RemoveChild(node);
                }
            }
        }
        internal void DeleteTopNode()
        {
            TopNode.ParentNode.RemoveChild(TopNode);
        }
        internal void SetXmlNodeDouble(string path, double? d, bool allowNegative)
        {
            SetXmlNodeDouble(path, d, null, "", allowNegative);
        }
        internal void SetXmlNodeDouble(string path, double? d, CultureInfo ci = null, string suffix="", bool allowNegative=true)
        {
            if (d.HasValue==false)
            {
                DeleteNode(path);
            }
            else
            {
                if (allowNegative==false && d.Value<0)
                {
                    throw new InvalidOperationException("Value can't be negative");
                }
                SetXmlNodeString(TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture) + suffix);
            }
        }
        internal void SetXmlNodeInt(string path, int? d, CultureInfo ci = null, bool allowNegative = true)
        {
            if (d == null)
            {
                DeleteNode(path);
            }   
            else
            {
                if(allowNegative==false && d.Value<0)
                {
                    throw new ArgumentException("Negative value not permitted");
                }
                SetXmlNodeString(TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture));
            }
        }
        internal void SetXmlNodeLong(string path, long? d, CultureInfo ci = null, bool allowNegative = true)
        {
            if (d == null)
            {
                DeleteNode(path);
            }
            else
            {
                if (allowNegative == false && d.Value < 0)
                {
                    throw new ArgumentException("Negative value not permitted");
                }
                SetXmlNodeString(TopNode, path, d.Value.ToString(ci ?? CultureInfo.InvariantCulture));
            }
        }
        readonly char[] _whiteSpaces = new char[] { '\t', '\n', '\r', ' ' };
        internal void SetXmlNodeStringPreserveWhiteSpace(string path, string value, bool removeIfBlank=false, bool insertFirst=false)
        {
            SetXmlNodeString(TopNode, path, value, removeIfBlank, insertFirst);
            if (value!=null &&  value.Length>0)
            {
                if(_whiteSpaces.Contains(value[0]) ||
                   _whiteSpaces.Contains(value[value.Length - 1]))
                {
                    var workNode = GetNode(path);
                    if(workNode.NodeType==XmlNodeType.Attribute)
                    {
                        workNode=workNode.ParentNode;
                    }
                    if(workNode.NodeType == XmlNodeType.Element)
                    {
                        ((XmlElement)workNode).SetAttribute("xml:space", "preserve");
                    }
                }
            }
        }

        internal void SetXmlNodeString(string path, string value)
        {
            SetXmlNodeString(TopNode, path, value, false, false);
        }
        internal void SetXmlNodeString(string path, string value, bool removeIfBlank)
        {
            SetXmlNodeString(TopNode, path, value, removeIfBlank, false);
        }
        internal void SetXmlNodeString(XmlNode node, string path, string value)
        {
            SetXmlNodeString(node, path, value, false, false);
        }
        internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank)
        {
            SetXmlNodeString(node, path, value, removeIfBlank, false);
        }
        internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
        {
            if (node == null)
            {
                return;
            }
            if (value == "" && removeIfBlank)
            {
                DeleteAllNode(path);
            }
            else
            {
                XmlNode nameNode = node.SelectSingleNode(path, NameSpaceManager);
                if (nameNode == null)
                {
                    CreateNode(path, insertFirst);
                    nameNode = node.SelectSingleNode(path, NameSpaceManager);
                }
                //if (nameNode.InnerText != value) HasChanged();
                nameNode.InnerText = value;
            }
        }
        internal void SetXmlNodeBool(string path, bool value)
        {
            SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
        }
        internal void SetXmlNodeBoolVml(string path, bool value)
        {
            SetXmlNodeString(TopNode, path, value ? "t" : "f", false, false);
        }

        internal void SetXmlNodeBool(string path, bool value, bool removeIf)
        {
            if (value == removeIf)
            {
                var node = TopNode.SelectSingleNode(path, NameSpaceManager);
                if (node != null)
                {
                    if (node is XmlAttribute attrib)
                    {
                        var elem = attrib.OwnerElement;
                        elem.RemoveAttribute(node.Name);
                    }
                    else
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
            else
            {
                SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
            }
        }
        internal void SetXmlNodePercentage(string path, double? value, bool allowNegative = true, double minMaxValue = 100D)
        {
            if (value.HasValue)
            {
                if (allowNegative == false && value < 0) throw (new ArgumentException("Negative percentage not allowed"));
                if (value < -minMaxValue || value > minMaxValue) throw (new ArgumentOutOfRangeException("value", $"Percentage out of range. Ranges from {(allowNegative ? 0 : -minMaxValue)}% to {minMaxValue}%"));
                SetXmlNodeString(path, ((int)(value.Value * 1000)).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                DeleteNode(path);
            }
        }
        internal void SetXmlNodeAngel(string path, double? value, string parameter = null, int minValue = 0, int maxValue = 360)
        {
            if (value.HasValue)
            {
                int v;
                if (!string.IsNullOrEmpty(parameter) && (value < minValue || value > maxValue))
                {
                    throw (new ArgumentOutOfRangeException(parameter, $"Value must be between {minValue} and {maxValue}"));
                }
                v = (int)(value * 60000);
                SetXmlNodeString(path, v.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                DeleteNode(path);
            }
        }
        internal void SetXmlNodeEmuToPt(string path, double? value)
        {
            if (value.HasValue)
            {
                int v;
                v = (int)(value * Drawing.ExcelDrawing.EMU_PER_POINT);
                SetXmlNodeString(path, v.ToString());
            }
            else
            {
                DeleteNode(path);
            }
        }
        internal void SetXmlNodeFontSize(string path, double? value, string propertyName, bool AllowNegative = true)
        {
            if (value.HasValue)
            {
                if (AllowNegative)
                {
                    if (value < 0 || value > 4000) throw (new ArgumentOutOfRangeException(propertyName, "Fontsize must be between 0 and 4000"));
                }
                else
                {
                    if (value < -4000 || value > 4000) throw (new ArgumentOutOfRangeException(propertyName, "Fontsize must be between -4000 and 4000"));
                }
                SetXmlNodeString(path, ((double)value * 100).ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                DeleteNode(path);
            }
        }
        internal bool ExistsNode(string path)
        {
            if (TopNode == null || TopNode.SelectSingleNode(path, NameSpaceManager) == null)
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
            if (node == null || node.SelectSingleNode(path, NameSpaceManager) == null)
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
            var value = GetXmlNodeString(path);
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            return GetXmlNodeBool(path);
        }
        internal bool? GetXmlNodeBoolNullableWithVal(string path)
        {
            var node = GetNode(path);
            if (node==null)
            {
                return null;
            }
            var value = node.Attributes["val"];
            if (value==null)
            {
                return true;
            }       
            else
            {
                return value.Value == "1" || value.Value == "-1" || value.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase);
            }
        }
        internal bool GetXmlNodeBool(string path)
        {
            return GetXmlNodeBool(path, false);
        }
        internal bool GetXmlNodeBool(string path, bool blankValue)
        {
            string value = GetXmlNodeString(path);
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
        internal static bool GetBoolFromString(string s)
        {
            return s != null && (s == "1" || s == "-1" || s.Equals("true", StringComparison.OrdinalIgnoreCase));
        }

        internal int GetXmlNodeInt(string path, int defaultValue=int.MinValue)
        {
            int i;
            if (int.TryParse(GetXmlNodeString(path), NumberStyles.Number, CultureInfo.InvariantCulture, out i))
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
            int a = GetXmlNodeInt(path);
            if (a < 0) return defaultValue;
            return a / 60000D;
        }
        internal double GetXmlNodeEmuToPt(string path)
        {
            var v = GetXmlNodeLong(path);
            if (v < 0) return 0;
            return (double)(v / (double)Drawing.ExcelDrawing.EMU_PER_POINT);
        }
        internal double? GetXmlNodeEmuToPtNull(string path)
        {
            var v = GetXmlNodeLongNull(path);
            if (v == null) return null;
            return (double)(v / (double)Drawing.ExcelDrawing.EMU_PER_POINT);
        }
        internal int? GetXmlNodeIntNull(string path)
        {
            int i;
            string s = GetXmlNodeString(path);
            if (s != "" && int.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out i))
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
            long l;
            string s = GetXmlNodeString(path);
            if (s != "" && long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out l))
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
            long l;
            string s = GetXmlNodeString(path);
            if (s != "" && long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out l))
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
            decimal d;
            if (decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
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
            decimal d;
            if (decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
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
            string s = GetXmlNodeString(path);
            if (s == "")
            {
                return null;
            }
            else
            {
                double v;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
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
            string s = GetXmlNodeString(path);
            if (s == "")
            {
                return double.NaN;
            }
            else
            {
                double v;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
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

            XmlNode nameNode = node.SelectSingleNode(path, NameSpaceManager);

            if (nameNode != null)
            {
                if (nameNode.NodeType == XmlNodeType.Attribute)
                {
                    return nameNode.Value != null ? nameNode.Value : "";
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
        internal string GetXmlNodeString(string path)
        {
            return GetXmlNodeString(TopNode, path);
        }
        internal static Uri GetNewUri(Packaging.ZipPackage package, string sUri)
        {
            var id = 1;
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
        internal T? GetXmlEnumNull<T>(string path, T? defaultValue=null) where T : struct, Enum
        {
            var v = GetXmlNodeString(path);
            if(string.IsNullOrEmpty(v))
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
            var p = GetXmlNodeString(path);
            if (p.EndsWith("%"))
            {
                if (double.TryParse(p.Substring(0, p.Length - 1), out d))
                    return d;
                else
                {
                    return null;
                }
            }
            else
            {
                if (double.TryParse(p, out d))
                    return d / 1000;
                else
                    return null;
            }
        }
        internal double GetXmlNodeFontSize(string path)
        {
            return (GetXmlNodeDoubleNull(path) ?? 0) / 100;
        }
        internal void RenameNode(XmlNode node, string prefix, string newName, string[] allowedChildren=null)
        {
            
            var doc = node.OwnerDocument;
            var newNode = doc.CreateElement(prefix, newName, NameSpaceManager.LookupNamespace(prefix));
            while (TopNode.ChildNodes.Count > 0)
            {
                if (allowedChildren == null || allowedChildren.Contains(TopNode.ChildNodes[0].LocalName))
                    newNode.AppendChild(TopNode.ChildNodes[0]);
                else
                    TopNode.RemoveChild(TopNode.ChildNodes[0]);
            }
            TopNode.ParentNode.ReplaceChild(newNode, TopNode);
            TopNode = newNode;
        }
        /// <summary>
        /// Insert the new node before any of the nodes in the comma separeted list
        /// </summary>
        /// <param name="parentNode">Parent node</param>
        /// <param name="beforeNodes">comma separated list containing nodes to insert after. Left to right order</param>
        /// <param name="newNode">The new node to be inserterd</param>
        internal void InserAfter(XmlNode parentNode, string beforeNodes, XmlNode newNode)
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
            if(insertAfter==null)
            {
                parentNode.AppendChild(newNode);
            }
            else
            {
                parentNode.InsertAfter(newNode, insertAfter);
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
            using (var stream = RecyclableMemory.GetStream(encoding.GetBytes(xml)))
            {
                LoadXmlSafe(xmlDoc, stream);
            }
        }
        internal void CreatespPrNode(string nodePath = "c:spPr", bool withLine = true)
        {
            if (!ExistsNode(nodePath))
            {
                var node = CreateNode(nodePath);
                if (withLine)
                    node.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/>";
                else
                    node.InnerXml = "<a:noFill/><a:effectLst/><a:sp3d/>";

            }
        }

        internal XmlNode GetOrCreateExtLstSubNode(string uriGuid, string prefix, string[] uriOrder=null)
        {
            foreach(XmlElement node in GetNodes("d:extLst/d:ext"))
            {
                if(node.Attributes["uri"].Value.Equals(uriGuid, StringComparison.OrdinalIgnoreCase))
                {
                    return node;
                }
            }
            var extLst = (XmlElement)CreateNode("d:extLst");
            XmlElement prependChild = null;
            if (uriOrder != null)
            {
                foreach (var child in extLst.ChildNodes)
                {
                    if (child is XmlElement e)
                    {
                        var uo1 = Array.IndexOf(uriOrder, e.GetAttribute("uri"));
                        var uo2 = Array.IndexOf(uriOrder, uriGuid);
                        if (uo1 > uo2)
                        {
                            prependChild = e;
                        }
                    }
                }
            }
            var newExt = TopNode.OwnerDocument.CreateElement("ext", ExcelPackage.schemaMain);
            if (!string.IsNullOrEmpty(prefix))
            {
                newExt.SetAttribute($"xmlns:{prefix}", NameSpaceManager.LookupNamespace(prefix));
            }

            newExt.SetAttribute("uri", uriGuid);
            if(prependChild==null)
            {
                extLst.AppendChild(newExt);
            }
            else
            {
                extLst.InsertBefore(newExt, prependChild);
            }

            return newExt;
        }
    }
}
