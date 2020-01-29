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
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A fill in a differential formatting record
    /// </summary>
    public class ExcelDxfFill : DxfStyleBase<ExcelDxfFill>
    {
        internal ExcelDxfFill(ExcelStyles styles)
            : base(styles)
        {
            PatternColor = new ExcelDxfColor(styles);
            BackgroundColor = new ExcelDxfColor(styles);
        }
        /// <summary>
        /// The pattern tyle
        /// </summary>
        public ExcelFillStyle? PatternType { get; set; }
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelDxfColor PatternColor { get; internal set; }
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelDxfColor BackgroundColor { get; internal set; }
        /// <summary>
        /// The Id
        /// </summary>
        protected internal override string Id
        {
            get
            {
                return GetAsString(PatternType) + "|" + (PatternColor == null ? "" : PatternColor.Id) + "|" + (BackgroundColor == null ? "" : BackgroundColor.Id);
            }
        }
        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            helper.CreateNode(path);
            SetValueEnum(helper, path + "/d:patternFill/@patternType", PatternType);
            SetValueColor(helper, path + "/d:patternFill/d:fgColor", PatternColor);
            SetValueColor(helper, path + "/d:patternFill/d:bgColor", BackgroundColor);
        }
        /// <summary>
        /// If the object has a value
        /// </summary>
        protected internal override bool HasValue
        {
            get 
            {
                return PatternType != null ||
                    PatternColor.HasValue ||
                    BackgroundColor.HasValue;
            }
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override ExcelDxfFill Clone()
        {
            return new ExcelDxfFill(_styles) {PatternType=PatternType, PatternColor=PatternColor.Clone(), BackgroundColor=BackgroundColor.Clone()};
        }
    }
}
