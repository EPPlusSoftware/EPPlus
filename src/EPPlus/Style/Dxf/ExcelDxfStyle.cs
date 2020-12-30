/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
 Date               Author                       Change
 *************************************************************************************************
 12/28/2020         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public abstract class ExcelDxfStyle : DxfStyleBase 
    {
        internal XmlHelperInstance _helper;
        internal ExcelDxfStyle(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles) : base(styles)
        {
            NumberFormat = new ExcelDxfNumberFormat(_styles);
            Border = new ExcelDxfBorderBase(_styles);
            Fill = new ExcelDxfFill(_styles);

            if (topNode != null)
            {
                _helper = new XmlHelperInstance(nameSpaceManager, topNode);
                if (_helper.ExistNode("d:numFmt"))
                {
                    NumberFormat.NumFmtID = _helper.GetXmlNodeInt("d:numFmt/@numFmtId");
                    NumberFormat.Format = _helper.GetXmlNodeString("d:numFmt/@formatCode");
                    if (NumberFormat.NumFmtID < 164 && string.IsNullOrEmpty(NumberFormat.Format))
                    {
                        NumberFormat.Format = ExcelNumberFormat.GetFromBuildInFromID(NumberFormat.NumFmtID);
                    }
                }
                if (_helper.ExistNode("d:border"))
                {
                    Border.Left = GetBorderItem(_helper, "d:border/d:left");
                    Border.Right = GetBorderItem(_helper, "d:border/d:right");
                    Border.Bottom = GetBorderItem(_helper, "d:border/d:bottom");
                    Border.Top = GetBorderItem(_helper, "d:border/d:top");
                }

                if (_helper.ExistNode("d:fill"))
                {
                    Fill.PatternType = GetPatternTypeEnum(_helper.GetXmlNodeString("d:fill/d:patternFill/@patternType"));
                    Fill.BackgroundColor = GetColor(_helper, "d:fill/d:patternFill/d:bgColor/");
                    Fill.PatternColor = GetColor(_helper, "d:fill/d:patternFill/d:fgColor/");
                }
            }
            else
            {
                _helper = new XmlHelperInstance(nameSpaceManager);
            }
            _helper.SchemaNodeOrder = new string[] { "font", "numFmt", "fill", "border" };
        }
        private ExcelDxfBorderItem GetBorderItem(XmlHelperInstance helper, string path)
        {
            ExcelDxfBorderItem bi = new ExcelDxfBorderItem(_styles);
            bi.Style = GetBorderStyleEnum(helper.GetXmlNodeString(path + "/@style"));
            bi.Color = GetColor(helper, path + "/d:color");
            return bi;
        }
        private static ExcelBorderStyle GetBorderStyleEnum(string style)
        {
            if (style == "") return ExcelBorderStyle.None;
            string sInStyle = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
            try
            {
                return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
            }
            catch
            {
                return ExcelBorderStyle.None;
            }

        }
        internal static ExcelFillStyle GetPatternTypeEnum(string patternType)
        {
            if (patternType == "") return ExcelFillStyle.None;
            patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);
            try
            {
                return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
            }
            catch
            {
                return ExcelFillStyle.None;
            }
        }

        internal virtual int DxfId { get; set; }
        /// <summary>
        /// Numberformat formatting settings
        /// </summary>
        public ExcelDxfNumberFormat NumberFormat { get; set; }
        /// <summary>
        /// Fill formatting settings
        /// </summary>
        public ExcelDxfFill Fill { get; set; }
        /// <summary>
        /// Border formatting settings
        /// </summary>
        public ExcelDxfBorderBase Border { get; set; }
        /// <summary>
        /// Id
        /// </summary>
        protected internal override string Id
        {
            get
            {
                return NumberFormat.Id + Border.Id + Fill.Id +
                    (AllowChange ? "" : DxfId.ToString());
            }
        }
        
        /// <summary>
        /// Creates the node
        /// </summary>
        /// <param name="helper">The helper</param>
        /// <param name="path">The XPath</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (NumberFormat.HasValue) NumberFormat.CreateNodes(helper, "d:numFmt");
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
        }
        /// <summary>
        /// If the object has a value
        /// </summary>
        protected internal override bool HasValue
        {
            get 
            {
                return  NumberFormat.HasValue || Fill.HasValue || Border.HasValue; 
            }
        }

    }
}
