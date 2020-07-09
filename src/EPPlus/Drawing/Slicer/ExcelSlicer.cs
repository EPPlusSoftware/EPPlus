/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.3
 ******0*******************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    public abstract class ExcelSlicer<T> : ExcelDrawing where T : ExcelSlicerCache
    {
        internal ExcelWorksheet _ws;
        protected XmlHelper _slicerXmlHelper;
        internal ExcelSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent=null) :
            base(drawings, node, "mc:AlternateContent/mc:Choice/xdr:graphicFrame", "xdr:nvGraphicFramePr/xdr:cNvPr", parent)
        {
            _ws = drawings.Worksheet;
        }
        internal ExcelSlicer(ExcelDrawings drawings, XmlNode node, XmlDocument slicerXml, ExcelGroupShape parent = null) :
            base(drawings, node, "mc:AlternateContent/mc:Choice/xdr:graphicFrame", "xdr:nvGraphicFramePr/xdr:cNvPr", parent)
        {
            _ws = drawings.Worksheet;
        }

        /// <summary>
        /// The caption text of the slicer.
        /// </summary>
        public string Caption 
        {
            get
            {
                return _slicerXmlHelper.GetXmlNodeString("@caption");
            }
            set
            {
                _slicerXmlHelper.SetXmlNodeString("@caption", value);
            }
        }
        /// <summary>
        /// Row height in points
        /// </summary>
        public double RowHeight 
        { 
            get
            {
                return _slicerXmlHelper.GetXmlNodeEmuToPt("@rowHeight");
            }
            set
            {
                _slicerXmlHelper.SetXmlNodeEmuToPt("@rowHeight", value);
            }
        }
        /// <summary>
        /// The start item
        /// </summary>
        public int StartItem
        {
            get
            { 
                return _slicerXmlHelper.GetXmlNodeInt("@startItem", 0);
            }
            set
            {
                _slicerXmlHelper.SetXmlNodeInt("@startItem", value, null, false);
            }
        }
        /// <summary>
        /// Number of columns. Default is 1.
        /// </summary>
        public int ColumnCount
        {
            get
            {
                return _slicerXmlHelper.GetXmlNodeInt("@columnCount", 1);
            }
            internal set
            {
                _slicerXmlHelper.SetXmlNodeInt("@columnCount", value, null, false);
            }
        }
        /// <summary>
        /// If the slicer view is locked or not.
        /// </summary>
        public bool LockedPosition
        {
            get
            {
                return _slicerXmlHelper.GetXmlNodeBool("@lockedPosition", false);
            }
            internal set
            {
                _slicerXmlHelper.SetXmlNodeBool("@lockedPosition", value, false);
            }
        }
        /// <summary>
        /// The build in slicer style.
        /// If set to Custom, the name in the <see cref="StyleName" /> is used 
        /// </summary>
        public eSlicerStyle Style
        {
            get
            {
                return StyleName.TranslateSlicerStyle();
            }
            set
            {
                if(value==eSlicerStyle.None)
                {
                    StyleName = "";
                }
                else if(value != eSlicerStyle.Custom)
                {
                    StyleName = "SlicerStyle" + value.ToString();
                }
            }
        }
        /// <summary>
        /// The style name used for the slicer.
        /// <seealso cref="Style"/>
        /// </summary>
        public string StyleName
        {
            get
            {
                return GetXmlNodeString("@style");
            }
            set
            {
                if(String.IsNullOrEmpty(value))
                {
                    Style = eSlicerStyle.None;
                }
                if(value.StartsWith("SlicerStyle", StringComparison.OrdinalIgnoreCase))
                {
                    var style = value.Substring(11).ToEnum<eSlicerStyle>(eSlicerStyle.Custom);
                    if(style!=eSlicerStyle.Custom || style!=eSlicerStyle.None)
                    {
                        Style = style;
                        return;
                    }
                }
                Style = eSlicerStyle.Custom;
                SetXmlNodeString("@style", value);
            }
        }
        internal string CacheName
        {
            get
            {
                return _slicerXmlHelper.GetXmlNodeString("@cache");
            }
            set
            {
                _slicerXmlHelper.SetXmlNodeString("@cache", value);
            }
        }
        protected internal ExcelSlicerCache _cache = null;
        public T Cache
        {
            get
            {
                if(_cache==null)
                {
                    _cache = _drawings.Worksheet.Workbook.GetSlicerCaches(CacheName);
                }
                return _cache as T;
            }
        }
    }
}
