/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/29/2023         EPPlus Software AB       Added
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Represents a cell alignment properties used for differential style formatting.
    /// </summary>
    public class ExcelDxfAlignment : DxfStyleBase
    {
        internal ExcelDxfAlignment(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
        }

        /// <summary>
        /// Horizontal alignment.
        /// </summary>
        public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
        /// <summary>
        /// Vertical alignment.
        /// </summary>
        public ExcelVerticalAlignment? VerticalAlignment { get; set; }
        int? _textRotation=null;
        /// <summary>
        /// String orientation in degrees. Values range from 0 to 180 or 255. 
        /// Setting the rotation to 255 will align text vertically.
        /// </summary>
        public int? TextRotation 
        { 
            get
            {
                return _textRotation;
            }
            set
            {
                if (value.HasValue && ((value < 0 || value > 180) && value != 255))
                {
                    throw new ArgumentOutOfRangeException("TextRotation out of range.");
                }
                _textRotation = value;
            }
        }

        /// <summary>
        /// Wrap the text
        /// </summary>
        public bool? WrapText { get; set; }
        int? _indent;
        /// <summary>
        /// The margin between the border and the text
        /// </summary>
        public int? Indent
        {
            get
            {
                return _indent;
            }
            set
            {
                if (value < 0 || value > 250)
                {
                    throw (new ArgumentOutOfRangeException("Indent must be between 0 and 250"));
                }
                _indent = value;
            }

        }
        /// <summary>
        /// The additional number of spaces of indentation to adjust for text in a cell.
        /// </summary>
        public int? RelativeIndent { get; set; }
        /// <summary>
        /// If the cells justified or distributed alignment should be used on the last line of text.
        /// </summary>
        public bool? JustifyLastLine { get; set; }
        /// <summary>
        /// Shrink the text to fit
        /// </summary>
        public bool? ShrinkToFit { get; set; }
        /// <summary>
        /// Reading order
        /// 0 - Context Dependent 
        /// 1 - Left-to-Right
        /// 2 - Right-to-Left
        /// </summary>
        public int? ReadingOrder { get; set; }
        /// <summary>
        /// Makes the text vertically. This is the same as setting <see cref="TextRotation"/> to 255.
        /// </summary>
        public void SetTextVertical()
        {
            TextRotation = 255;
        }
        /// <summary>
        /// If the dxf style has any values set.
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return HorizontalAlignment.HasValue || 
                       VerticalAlignment.HasValue || 
                       TextRotation.HasValue ||
                       WrapText.HasValue ||
                       Indent.HasValue ||
                       RelativeIndent.HasValue ||
                       JustifyLastLine.HasValue ||
                       ShrinkToFit.HasValue ||
                       ReadingOrder.HasValue;
            }
        }

        internal override string Id 
        {
            get
            {
                return GetAsString(HorizontalAlignment) + "|" + 
                       GetAsString(VerticalAlignment) + "|" +
                       GetAsString(TextRotation) + "|" +
                       GetAsString(WrapText) + "|" +
                       GetAsString(Indent) + "|" +
                       GetAsString(RelativeIndent) + "|" +
                       GetAsString(JustifyLastLine) + "|" +
                       GetAsString(ShrinkToFit) + "|" +
                       GetAsString(ReadingOrder);

            }
        }
        internal static string GetEmptyId()
        {
            return "||||||||";
		}
		/// <summary>
		/// Clears all properties
		/// </summary>
		public override void Clear()
        {

            HorizontalAlignment = null;
            VerticalAlignment = null;
            TextRotation = null;
            WrapText = null;
            Indent = null;
            RelativeIndent = null;
            JustifyLastLine = null;
            ShrinkToFit = null;
            ReadingOrder = null;
        }

        internal override DxfStyleBase Clone()
        {
            return new ExcelDxfAlignment(_styles, _callback) 
            { 
                HorizontalAlignment = HorizontalAlignment,
                VerticalAlignment = VerticalAlignment,
                TextRotation = TextRotation,
                WrapText = WrapText,
                Indent = Indent,
                RelativeIndent = RelativeIndent,
                JustifyLastLine = JustifyLastLine,
                ShrinkToFit = ShrinkToFit,
                ReadingOrder = ReadingOrder
            };
        }

        internal override void CreateNodes(XmlHelper helper, string path)
        {
            SetValueEnum(helper, path + "/@horizontal", HorizontalAlignment);
            SetValueEnum(helper, path + "/@vertical", VerticalAlignment);
            SetValue(helper, path + "/@textRotation", TextRotation);
            SetValueBool(helper, path + "/@wrapText", WrapText);
            SetValue(helper, path + "/@indent", Indent);
            SetValue(helper, path + "/@relativeIndent", RelativeIndent);
            SetValueBool(helper, path + "/@justifyLastLine", JustifyLastLine);
            SetValueBool(helper, path + "/@shrinkToFit", ShrinkToFit);
            SetValue(helper, path + "/@readingOrder", ReadingOrder);
        }

        internal override void SetStyle()
        {
            if(_callback!=null)
            {
                _callback.Invoke(eStyleClass.Style, eStyleProperty.HorizontalAlign, HorizontalAlignment);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.VerticalAlign, VerticalAlignment);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.TextRotation, TextRotation);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.WrapText, WrapText);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.Indent, Indent);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.JustifyLastLine, JustifyLastLine);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.ShrinkToFit, ShrinkToFit);
                _callback.Invoke(eStyleClass.Style, eStyleProperty.ReadingOrder, ReadingOrder);

            }
        }
        internal override void SetValuesFromXml(XmlHelper helper)
        {
            HorizontalAlignment = GetEnumValue<ExcelHorizontalAlignment>(helper.GetXmlNodeString("d:alignment/@horizontal"));
            VerticalAlignment = GetEnumValue<ExcelVerticalAlignment>(helper.GetXmlNodeString("d:alignment/@vertical"));
            TextRotation = helper.GetXmlNodeIntNull("d:alignment/@textRotation");
            WrapText = helper.GetXmlNodeBoolNullable("d:alignment/@wrapText");
            Indent = helper.GetXmlNodeIntNull("d:alignment/@indent");
            RelativeIndent = helper.GetXmlNodeIntNull("d:alignment/@relativeIndent");
            JustifyLastLine = helper.GetXmlNodeBoolNullable("d:alignment/@justifyLastLine");
            ShrinkToFit = helper.GetXmlNodeBoolNullable("d:alignment/@shrinkToFit");
            ReadingOrder = helper.GetXmlNodeIntNull("d:alignment/@readingOrder");            
        }

        private T? GetEnumValue<T>(string v) where T : struct
        {
            if (v == null)
            {
                return default;
            }
            else
            {
                return v.ToEnum<T>();
            }

        }
    }
}
