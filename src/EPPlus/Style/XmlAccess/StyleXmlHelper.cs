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
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml helper class for cell style classes
    /// </summary>
    public abstract class  StyleXmlHelper : XmlHelper
    {
        internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        { 

        }
        internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }
        internal abstract XmlNode CreateXmlNode(XmlNode top);
        internal abstract string Id
        {
            get;
        }
        internal long useCnt=0;
        internal int newID=int.MinValue;
        internal bool GetBoolValue(XmlNode topNode, string path)
        {
            var node = topNode.SelectSingleNode(path, NameSpaceManager);
            if (node is XmlAttribute)
            {
                return node.Value != "0";
            }
            else
            {
                if (node != null && ((node.Attributes["val"] != null && node.Attributes["val"].Value != "0") || node.Attributes["val"] == null))
                {
                    return true;
                }
                else
                {
                    return false;
                }                
            }
        }

    }
}
