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
    /// <summary>
    /// Base class for differential formatting styles
    /// </summary>
    public abstract class ExcelDxfStyleBase : DxfStyleBase                                         
    {
        internal XmlHelperInstance _helper;            
        //internal protected string _dxfIdPath;

        internal ExcelDxfStyleBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {
            //_dxfIdPath = dxfIdPath;
            Border = new ExcelDxfBorderBase(_styles, callback);
            Fill = new ExcelDxfFill(_styles, callback);

            if (topNode != null)
            {
                _helper = new XmlHelperInstance(nameSpaceManager, topNode);
                Border.SetValuesFromXml(_helper);
                Fill.SetValuesFromXml(_helper);
            }
            else
            {
                _helper = new XmlHelperInstance(nameSpaceManager);
            }
            _helper.SchemaNodeOrder = new string[] { "font", "numFmt", "fill", "border" };
        }
        internal virtual int DxfId { get; set; } = int.MinValue;
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
        internal override string Id
        {
            get
            {
                return GetId() + ExcelDxfFont.GetEmptyId() + ExcelDxfNumberFormat.GetEmptyId() + ExcelDxfAlignment.GetEmptyId() + ExcelDxfProtection.GetEmptyId();
            }
        }
        internal virtual string GetId()
        {
			return Border.GetId() + Fill.Id;
		}
		/// <summary>
		/// Creates the node
		/// </summary>
		/// <param name="helper">The helper</param>
		/// <param name="path">The XPath</param>
		internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (Fill.HasValue) Fill.CreateNodes(helper, "d:fill");
            if (Border.HasValue) Border.CreateNodes(helper, "d:border");
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                Border.SetStyle();
                Fill.SetStyle();
            }
        }

        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get 
            {
                return  Fill.HasValue || Border.HasValue; 
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            Fill.Clear();
            Border.Clear();
        }
        internal ExcelDxfStyle ToDxfStyle()
        {
            if (this is ExcelDxfStyle s)
            {
                return s;
            }
            else
            {
                var ns = new ExcelDxfStyle(_styles.NameSpaceManager, null, _styles, null)
                {
                    Border = Border,
                    Fill = Fill,
                    DxfId = DxfId,
                    Font = new ExcelDxfFont(_styles, _callback),
                    NumberFormat = new ExcelDxfNumberFormat(_styles, _callback),
                    _helper = _helper
                };
                ns.Font.SetValuesFromXml(_helper);
                return ns;
            }
        }
        internal ExcelDxfSlicerStyle ToDxfSlicerStyle()
        {
            if (this is ExcelDxfSlicerStyle s)
            {
                return s;
            }
            else
            {
                var ns = new ExcelDxfSlicerStyle(_styles.NameSpaceManager, null, _styles, null)
                {
                    Border = Border,
                    Fill = Fill,
                    DxfId = DxfId,
                    Font = new ExcelDxfFont(_styles, _callback),
                    _helper = _helper
                };
                ns.Font.SetValuesFromXml(_helper);
                return ns;
            }
        }
        internal ExcelDxfTableStyle ToDxfTableStyle()
        {
            if(this is ExcelDxfTableStyle s)
            {
                return s;
            }
            else
            {
                var ns = new ExcelDxfTableStyle(_styles.NameSpaceManager, null, _styles)
                {
                    Border = Border,
                    Fill = Fill,
                    DxfId = DxfId,
                    Font = new ExcelDxfFont(_styles, _callback),
                    _helper = _helper
                };
                ns.Font.SetValuesFromXml(_helper);
                return ns;
            }
        }
        internal ExcelDxfStyleLimitedFont ToDxfLimitedStyle()
        {
            if (this is ExcelDxfStyleLimitedFont s)
            {
                return s;
            }
            else
            {
                var ns = new ExcelDxfStyleLimitedFont(_styles.NameSpaceManager, null, _styles, _callback)
                {
                    Border = Border,
                    Fill = Fill,
                    DxfId = DxfId,
                    Font = new ExcelDxfFontBase(_styles,_callback),
                    _helper = _helper
                };
                ns.Font.SetValuesFromXml(_helper);
                return ns;
            }
        }

		internal ExcelDxfStyleConditionalFormatting ToDxfConditionalFormattingStyle()
        {
            if (this is ExcelDxfStyleConditionalFormatting s)
            {
                return s;
            }
            else
            {
                var ns = new ExcelDxfStyleConditionalFormatting(_styles.NameSpaceManager, null, _styles, _callback)
                {
                    Border = Border,
                    Fill = Fill,
                    NumberFormat = new ExcelDxfNumberFormat(_styles, _callback),
                    DxfId = DxfId,
                    Font = new ExcelDxfFontBase(_styles, _callback),
                    _helper = _helper
                };
                ns.NumberFormat.SetValuesFromXml(_helper);
                ns.Font.SetValuesFromXml(_helper);
                return ns;
            }
        }
    }
}
