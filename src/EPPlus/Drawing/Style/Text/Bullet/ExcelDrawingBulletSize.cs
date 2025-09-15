/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingBulletSize : XmlHelper
    {
        const string BuSzPct = "/a:buSzPct";
        const string BuSzPts = "/a:buSzPts";
        const string BuSzTx = "/a:buSzTx";
        string _path;
        XmlNode _currentNode=null;
        internal ExcelDrawingBulletSize(XmlNamespaceManager nsm, XmlNode topNode, string path, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _path = path;
            var node = GetNode(path);
            if(node!=null)
            {
                _currentNode = GetNode(path + BuSzPct);
                if (_currentNode == null)
                {
                    _currentNode = GetNode(path + BuSzPts);
                    if (_currentNode == null)
                    {
                        Type = eBulletSizeType.FollowText;
                    }
                    else
                    {
                        Type = eBulletSizeType.Points;
                    }
                }
                else 
                {
                    Type = eBulletSizeType.PercentOfText;
                }
            }
            else
            {
                Type = eBulletSizeType.PercentOfText;
            }
        }
        public eBulletSizeType Type 
        { 
            get; 
        }
        /// <summary>
        /// The value if <see cref="Type"/> is set to PercentOfText or Points.
        /// </summary>
        public double? Value
        {
            get
            {
                if (Type == eBulletSizeType.PercentOfText)
                {
                    return XmlHelper.GetRichTextPropertyDouble(_currentNode) / 1000;
                }
                else if(Type == eBulletSizeType.Points)
                {
                    return XmlHelper.GetRichTextPropertyDouble(_currentNode) / 100;
                }
                return null;
            }
            internal set
            {
                if (value.HasValue == false)
                {
                    _currentNode.Attributes.Remove(_currentNode.Attributes["val"]); 
                }   
                if(Type == eBulletSizeType.PercentOfText)
                {
                    SetXmlNodeString(_currentNode, "/@val", (value.Value * 1000).ToString(CultureInfo.InvariantCulture));
                }
                else if (Type == eBulletSizeType.Points)
                {
                    SetXmlNodeString(_currentNode, "/@val", (value.Value * 100).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    throw new InvalidOperationException("Bullets with type eBulletSizeType.FollowText cannot have a Value");
                }
            }
        }
        /// <summary>
        /// Sets the bullet size to follow the text.
        /// </summary>
        public void SetFollowText()
        {
            if (Type == eBulletSizeType.FollowText) return;
            DeleteNode(_path + BuSzPts);
            DeleteNode(_path + BuSzPct);
            CreateNode(_path + BuSzTx);
        }
        /// <summary>
        /// Sets the bullet to a percentage of the text within the paragraph.
        /// </summary>
        /// <param name="value">The value in percent, where 100% is 100</param>
        public void SetPercent(double value)
        {
            DeleteNode(_path + BuSzPts);
            DeleteNode(_path + BuSzTx);
            CreateNode(_path + BuSzPct);

            Value = value * 100;
        }
        /// <summary>
        /// Sets the bullet to a size in points.
        /// </summary>
        /// <param name="value">Value in points</param>
        public void SetPoint(double value)
        {
            DeleteNode(_path + BuSzPct);
            DeleteNode(_path + BuSzTx);
            CreateNode(_path + BuSzPts);

            Value = value * 1000;
        }
    }
}