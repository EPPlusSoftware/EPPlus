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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD
{
    /// <summary>
    /// The points and vectors contained within the backdrop define a plane in 3D space
    /// </summary>
    public class ExcelDrawingScene3DBackDrop : XmlHelper
    {
        private readonly string _anchorPath = "{0}/a:anchor";
        private readonly string _normPath = "{0}/a:norm";
        private readonly string _upPath = "{0}/a:up";
        private readonly Action<bool> _initParent;
        internal ExcelDrawingScene3DBackDrop(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {           
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "anchor", "norm", "up"});

            _anchorPath = string.Format(_anchorPath, path);
            _normPath = string.Format(_normPath, path);
            _upPath = string.Format(_upPath, path);
            _initParent = initParent;
        }


        ExcelDrawingPoint3D _anchorPoint = null;
        /// <summary>
        /// The anchor point
        /// </summary>
        public ExcelDrawingPoint3D AnchorPoint
        {
            get
            {
                if (_anchorPoint == null)
                {
                    _anchorPoint = new ExcelDrawingPoint3D(NameSpaceManager, TopNode, SchemaNodeOrder, _anchorPath, "", InitXml);
                }
                return _anchorPoint;
            }
        }
        ExcelDrawingPoint3D _upVector = null;
        /// <summary>
        /// The up vector
        /// </summary>
        public ExcelDrawingPoint3D UpVector
        {
            get
            {
                if (_upVector == null)
                {
                    _upVector = new ExcelDrawingPoint3D(NameSpaceManager, TopNode, SchemaNodeOrder, _upPath, "d", InitXml);
                }
                return _upVector;

            }
        }
        ExcelDrawingPoint3D _normalVector = null;
        /// <summary>
        /// The normal vector
        /// </summary>
        public ExcelDrawingPoint3D NormalVector
        {
            get
            {
                if (_normalVector == null)
                {
                    _normalVector = new ExcelDrawingPoint3D(NameSpaceManager, TopNode, SchemaNodeOrder, _normPath, "d", InitXml);
                }
                return _normalVector;

            }
        }
        private void InitXml(bool delete)
        {
            _initParent(delete);
            AnchorPoint.InitXml();
            UpVector.InitXml();
            NormalVector.InitXml();
        }
    }
}
