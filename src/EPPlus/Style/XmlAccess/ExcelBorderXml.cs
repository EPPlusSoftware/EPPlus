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
    /// Xml access class for border top level
    /// </summary>
    public sealed class ExcelBorderXml : StyleXmlHelper
    {
        internal ExcelBorderXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {

        }
        internal ExcelBorderXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _left = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(leftPath, nsm));
            _right = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(rightPath, nsm));
            _top = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(topPath, nsm));
            _bottom = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(bottomPath, nsm));
            _diagonal = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(diagonalPath, nsm));
            _diagonalUp = GetBoolValue(topNode, diagonalUpPath);
            _diagonalDown = GetBoolValue(topNode, diagonalDownPath);
        }
        internal override string Id
        {
            get
            {
                return Left.Id + Right.Id + Top.Id + Bottom.Id + Diagonal.Id + DiagonalUp.ToString() + DiagonalDown.ToString();
            }
        }
        const string leftPath = "d:left";
        ExcelBorderItemXml _left = null;
        /// <summary>
        /// Left border style properties
        /// </summary>
        public ExcelBorderItemXml Left
        {
            get
            {
                return _left;
            }
            internal set
            {
                _left = value;
            }
        }
        const string rightPath = "d:right";
        ExcelBorderItemXml _right = null;
        /// <summary>
        /// Right border style properties
        /// </summary>
        public ExcelBorderItemXml Right
        {
            get
            {
                return _right;
            }
            internal set
            {
                _right = value;
            }
        }
        const string topPath = "d:top";
        ExcelBorderItemXml _top = null;
        /// <summary>
        /// Top border style properties
        /// </summary>
        public ExcelBorderItemXml Top
        {
            get
            {
                return _top;
            }
            internal set
            {
                _top = value;
            }
        }
        const string bottomPath = "d:bottom";
        ExcelBorderItemXml _bottom = null;
        /// <summary>
        /// Bottom border style properties
        /// </summary>
        public ExcelBorderItemXml Bottom
        {
            get
            {
                return _bottom;
            }
            internal set
            {
                _bottom = value;
            }
        }
        const string diagonalPath = "d:diagonal";
        ExcelBorderItemXml _diagonal = null;
        /// <summary>
        /// Diagonal border style properties
        /// </summary>
        public ExcelBorderItemXml Diagonal
        {
            get
            {
                return _diagonal;
            }
            internal set
            {
                _diagonal = value;
            }
        }
        const string diagonalUpPath = "@diagonalUp";
        bool _diagonalUp = false;
        /// <summary>
        /// Diagonal up border
        /// </summary>
        public bool DiagonalUp
        {
            get
            {
                return _diagonalUp;
            }
            internal set
            {
                _diagonalUp = value;
            }
        }
        const string diagonalDownPath = "@diagonalDown";
        bool _diagonalDown = false;
        /// <summary>
        /// Diagonal down border
        /// </summary>
        public bool DiagonalDown
        {
            get
            {
                return _diagonalDown;
            }
            internal set
            {
                _diagonalDown = value;
            }
        }

        internal ExcelBorderXml Copy()
        {
            ExcelBorderXml newBorder = new ExcelBorderXml(NameSpaceManager);
            newBorder.Bottom = _bottom.Copy();
            newBorder.Diagonal = _diagonal.Copy();
            newBorder.Left = _left.Copy();
            newBorder.Right = _right.Copy();
            newBorder.Top = _top.Copy();
            newBorder.DiagonalUp = _diagonalUp;
            newBorder.DiagonalDown = _diagonalDown;

            return newBorder;

        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            CreateNode(leftPath);
            topNode.AppendChild(_left.CreateXmlNode(TopNode.SelectSingleNode(leftPath, NameSpaceManager)));
            CreateNode(rightPath);
            topNode.AppendChild(_right.CreateXmlNode(TopNode.SelectSingleNode(rightPath, NameSpaceManager)));
            CreateNode(topPath);
            topNode.AppendChild(_top.CreateXmlNode(TopNode.SelectSingleNode(topPath, NameSpaceManager)));
            CreateNode(bottomPath);
            topNode.AppendChild(_bottom.CreateXmlNode(TopNode.SelectSingleNode(bottomPath, NameSpaceManager)));
            CreateNode(diagonalPath);
            topNode.AppendChild(_diagonal.CreateXmlNode(TopNode.SelectSingleNode(diagonalPath, NameSpaceManager)));
            if (_diagonalUp)
            {
                SetXmlNodeString(diagonalUpPath, "1");
            }
            if (_diagonalDown)
            {
                SetXmlNodeString(diagonalDownPath, "1");
            }
            return topNode;
        }
    }
}
