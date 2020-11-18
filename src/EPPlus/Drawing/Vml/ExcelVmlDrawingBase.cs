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

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Drawing object used for comments
    /// </summary>
    public class ExcelVmlDrawingBase : XmlHelper
    {
        internal ExcelVmlDrawingBase(XmlNode topNode, XmlNamespaceManager ns) :
            base(ns, topNode)
        {
            SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }
        /// <summary>
        /// The Id of the vml drawing
        /// </summary>
        public string Id 
        {
            get
            {
                return GetXmlNodeString("@id");
            }
            set
            {
                SetXmlNodeString("@id",value);
            }
        }
        /// <summary>
        /// The Id of the vml drawing
        /// </summary>
        public string SpId
        {
            get
            {
                return GetXmlNodeString("@o:spid");
            }
            set
            {
                SetXmlNodeString("@o:spid", value);
            }
        }
        /// <summary>
        /// Alternative text to be displayed instead of a graphic.
        /// </summary>
        public string AlternativeText
        {
            get
            {
                return GetXmlNodeString("@alt");
            }
            set
            {
                SetXmlNodeString("@alt", value);
            }
        }
        /// <summary>
        /// Anchor coordinates for the drawing
        /// </summary>
        internal string Anchor
        {
            get
            {
                return GetXmlNodeString("x:ClientData/x:Anchor");
            }
            set
            {
                SetXmlNodeString("x:ClientData/x:Anchor", value);
            }
        }

        #region "Style Handling methods"
        /// <summary>
        /// Gets a style from the semi-colo separated list with the specifik key
        /// </summary>
        /// <param name="style">The list</param>
        /// <param name="key">The key to search for</param>
        /// <param name="value">The value to return</param>
        /// <returns>True if found</returns>
        protected bool GetStyle(string style, string key, out string value)
        {
            string[]styles = style.Split(';');
            foreach(string s in styles)
            {
                if (s.IndexOf(':') > 0)
                {
                    string[] split = s.Split(':');
                    if (split[0] == key)
                    {
                        value=split[1];
                        return true;
                    }
                }
                else if (s == key)
                {
                    value="";
                    return true;
                }
            }
            value="";
            return false;
        }
        /// <summary>
        /// Sets the style in a semicolon separated list
        /// </summary>
        /// <param name="style">The list</param>
        /// <param name="key">The key</param>
        /// <param name="value">The value</param>
        /// <returns>The new list</returns>
        internal protected string SetStyle(string style, string key, string value)
        {
            string[] styles = style.Split(';');
            string newStyle="";
            bool changed = false;
            foreach (string s in styles)
            {
                if (!string.IsNullOrEmpty(s))
                {
                    string[] split = s.Split(':');
                    if (split[0].Trim() == key)
                    {
                        if (value.Trim() != "") //If blank remove the item
                        {
                            newStyle += key + ':' + value;
                        }
                        changed = true;
                    }
                    else
                    {
                        newStyle += s;
                    }
                    newStyle += ';';
                }
            }
            if (!changed)
            {
                newStyle += key + ':' + value;
            }
            else
            {
                newStyle = newStyle.Substring(0, newStyle.Length - 1);
            }
            return newStyle;
        }
        #endregion
    }
}
