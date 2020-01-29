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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A point in a 3D space
    /// </summary>
    public class ExcelDrawingPoint3D : XmlHelper
    {
        private readonly string _xPath = "{0}/@{1}x";
        private readonly string _yPath = "{0}/@{1}y";
        private readonly string _zPath = "{0}/@{1}z";
        private readonly Action<bool> _initParent;
        bool isInit = false;
        internal ExcelDrawingPoint3D(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path, string prefix, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _xPath = string.Format(_xPath, path, prefix);
            _yPath = string.Format(_yPath, path, prefix);
            _zPath = string.Format(_zPath, path, prefix);
            _initParent = initParent;
        }
        /// <summary>
        /// The X coordinate in point
        /// </summary>
        public double X
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_xPath) ?? 0;
            }
            set
            {
                if(isInit==false) _initParent(false);
                SetXmlNodeEmuToPt(_xPath, value);
            }
        }
        /// <summary>
        /// The Y coordinate
        /// </summary>
        public double Y
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_yPath) ?? 0;
            }
            set
            {
                if (isInit == false) _initParent(false);
                SetXmlNodeEmuToPt(_yPath, value);
            }
        }
        /// <summary>
        /// The Z coordinate
        /// </summary>
        public double Z
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_zPath) ?? 0;
            }
            set
            {
                if (isInit == false) _initParent(false);
                SetXmlNodeEmuToPt(_zPath, value);
            }
        }
        internal void InitXml()
        {
            if (isInit==false)
            {
                isInit = true;
                if (!ExistNode(_xPath))
                {
                    X = 0;
                    Y = 0;
                    Z = 0;
                }
            }
        }

    }
}
