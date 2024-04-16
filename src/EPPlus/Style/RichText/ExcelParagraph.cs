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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Handels paragraph text
    /// </summary>
    public sealed class ExcelParagraph : ExcelTextFont
    {
        internal ExcelParagraph(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager ns, XmlNode rootNode, string path, string[] schemaNodeOrder) : 
            base(pictureRelationDocument, ns, rootNode, path + "a:rPr", schemaNodeOrder)
        { 

        }
        const string TextPath = "../a:t";
        /// <summary>
        /// Text
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString(TextPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(TextPath, value);
            }
        }
        /// <summary>
        /// If the paragraph is the first in the collection
        /// </summary>
        public bool IsFirstInParagraph
        {
            get
            {
                var parent = _rootNode.ParentNode;
                for (int i=0;i<parent.ChildNodes.Count;i++)
                {
                    if (parent.ChildNodes[i].LocalName == "r")
                    {
                        return parent.ChildNodes[i] == _rootNode;
                    }
                }
                return false;
            }
        }
        /// <summary>
        /// If the paragraph is the last in the collection
        /// </summary>
        public bool IsLastInParagraph
        {
            get
            {
                var parent = _rootNode.ParentNode;
                for (int i = parent.ChildNodes.Count-1; i >=0 ; i--)
                {
                    if (parent.ChildNodes[i].LocalName == "r")
                    {
                        return parent.ChildNodes[i] == _rootNode;
                    }
                }
                return false;
            }
        }
    }
}
