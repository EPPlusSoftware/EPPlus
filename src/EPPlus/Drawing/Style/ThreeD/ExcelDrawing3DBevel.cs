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
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD
{
    /// <summary>
    /// Defines a bevel off a shape
    /// </summary>
    public class ExcelDrawing3DBevel : XmlHelper
    {
        bool _isInit = false;
        private string _path;
        private readonly string _widthPath = "{0}/@w";
        private readonly string _heightPath = "{0}/@h";
        private readonly string _typePath="{0}/@prst";
        private readonly Action<bool> _initParent;
        internal ExcelDrawing3DBevel(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _path = path;
            _widthPath = string.Format(_widthPath, path);
            _heightPath = string.Format(_heightPath, path);
            _typePath = string.Format(_typePath, path);
            _initParent = initParent;
        }
        /// <summary>
        /// The width of the bevel in points (pt)
        /// </summary>
        public double Width
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_widthPath) ?? 6;
            }
            set
            {
                if (!_isInit) InitXml();
                SetXmlNodeEmuToPt(_widthPath, value);
            }
        }

        private void InitXml()
        {
            if (_isInit==false)
            {
                _isInit = true;
                if (!ExistNode(_typePath))
                {
                    _initParent(false);
                    Height = 6;
                    Width = 6;
                    BevelType = eBevelPresetType.Circle;
                }
            }
        }

        /// <summary>
        /// The height of the bevel in points (pt)
        /// </summary>
        public double Height
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_heightPath) ?? 6;
            }
            set
            {
                if(!_isInit) InitXml();
                SetXmlNodeEmuToPt(_heightPath, value);
            }
        }
        /// <summary>
        /// A preset bevel that can be applied to a shape.
        /// </summary>
        public eBevelPresetType BevelType
        {
            get
            {
                return GetXmlNodeString(_typePath).ToEnum(eBevelPresetType.Circle);
            }
            set
            {
                if(value==eBevelPresetType.None)
                {
                    DeleteNode(_typePath);
                    DeleteNode(_heightPath);
                    DeleteNode(_widthPath);
                }
                else
                {
                    if (!_isInit) InitXml();
                    SetXmlNodeString(_typePath, value.ToEnumString());
                }
            }
        }
    }
}