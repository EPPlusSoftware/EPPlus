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

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A coordinate in 3D space.
    /// </summary>
    public class ExcelDrawingSphereCoordinate : XmlHelper
    {
        /// <summary>
        /// XPath 
        /// </summary>
        internal protected string _path;
        private readonly string _latPath ="{0}/@lat";
        private readonly string _lonPath = "{0}/@lon";
        private readonly string _revPath = "{0}/@rev";
        private readonly Action<bool> _initParent;
        internal ExcelDrawingSphereCoordinate(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string path, Action<bool> initParent) : base(nameSpaceManager, topNode)
        {
            _path = path;
            _latPath = string.Format(_latPath, path);
            _lonPath = string.Format(_lonPath, path);
            _revPath = string.Format(_revPath, path);
            _initParent = initParent;
        }
        /// <summary>
        /// The latitude value of the rotation
        /// </summary>
        public double Latitude
        {
            get
            {
                return GetXmlNodeAngle(_latPath);
            }
            set
            {
                InitXml();
                SetXmlNodeAngle(_latPath, value, "Latitude");
            }
        }
        /// <summary>
        /// The longitude value of the rotation
        /// </summary>
        public double Longitude
        {
            get
            {
                return GetXmlNodeAngle(_lonPath);
            }
            set
            {
                InitXml();
                SetXmlNodeAngle(_lonPath, value, "Longitude");
            }
        }
        /// <summary>
        /// The revolution around the central axis in the rotation
        /// </summary>
        public double Revolution
        {
            get
            {
                return GetXmlNodeAngle(_revPath);
            }
            set
            {
                InitXml();
                SetXmlNodeAngle(_revPath, value, "Revolution");
            }
        }
        bool isInit = false;
        /// <summary>
        /// All values are required, so init them on any set.
        /// </summary>
        private void InitXml()
        {
            if(isInit==false)
            {
                isInit = true;
                if (!ExistsNode(_latPath))
                {
                    _initParent(false);
                    Latitude = 0;
                    Longitude = 0;
                    Revolution = 0;
                }
            }
        }
    }
}
