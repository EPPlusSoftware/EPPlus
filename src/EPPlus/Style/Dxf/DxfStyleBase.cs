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
using OfficeOpenXml.Drawing;
using System;
using System.Drawing;
using System.Globalization;
namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Base class for differential formatting styles. 
    /// </summary>
    public abstract class DxfStyleBase
    {
        internal ExcelStyles _styles;
        internal Action<eStyleClass, eStyleProperty, object> _callback;
        internal DxfStyleBase(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        {
            _styles = styles;
            _callback = callback;
            AllowChange = false; //Don't touch this value in the styles.xml (by default). When Dxfs is fully implemented this can be removed.
        }
        /// <summary>
        /// Reset all properties for the style.
        /// </summary>
        public abstract void Clear();
        /// <summary>
        /// The id
        /// </summary>
        protected internal abstract string Id { get; }
        /// <summary>
        /// If the style has any value set
        /// </summary>
        public abstract bool HasValue{get;}
        /// <summary>
        /// Create the nodes
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The Xpath</param>
        protected internal abstract void CreateNodes(XmlHelper helper, string path);
        /// <summary>
        /// Sets the values from an XmlHelper instance. 
        /// </summary>
        /// <param name="helper">The helper</param>
        protected internal virtual void SetValuesFromXml(XmlHelper helper)
        {

        }
        /// <summary>
        /// Set the cell style values from the dxf using the callback method.
        /// </summary>
        internal abstract void SetStyle();

        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns></returns>
        protected internal abstract DxfStyleBase Clone();
        /// <summary>
        /// Set the color value
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The x path</param>
        /// <param name="color">The color</param>
        protected void SetValueColor(XmlHelper helper,string path, ExcelDxfColor color)
        {
            if (color != null && color.HasValue)
            {
                if (color.Color != null)
                {
                    SetValue(helper, path + "/@rgb", color.Color.Value.ToArgb().ToString("x"));
                }
                else if (color.Auto != null)
                {
                    SetValueBool(helper, path + "/@auto", color.Auto);
                }
                else if (color.Theme != null)
                {
                    SetValue(helper, path + "/@theme", (int)color.Theme);
                }
                else if (color.Index != null)
                {
                    SetValue(helper, path + "/@indexed", (int)color.Index);
                }
                if (color.Tint != null)
                {
                    SetValue(helper, path + "/@tint", color.Tint);
                }
            }
        }
        /// <summary>
        /// Same as SetValue but will set first char to lower case.
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The Xpath</param>
        /// <param name="v">The value</param>
        internal protected void SetValueEnum(XmlHelper helper, string path, Enum v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                var s = v.ToString();
                s = s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1);
                helper.SetXmlNodeString(path, s);
            }
        }
        /// <summary>
        /// Sets the value
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The x path</param>
        /// <param name="v">The object</param>
        internal protected void SetValue(XmlHelper helper, string path, object v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                if(v is double d)
                {
                    helper.SetXmlNodeString(path, d.ToString(CultureInfo.InvariantCulture));
                }
                else if (v is int i)
                {
                    helper.SetXmlNodeString(path, i.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    helper.SetXmlNodeString(path, v.ToString());
                }
            }
        }
        /// <summary>
        /// Sets the value
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The x path</param>
        /// <param name="v">The string</param>
        internal protected void SetValue(XmlHelper helper, string path, string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                helper.DeleteNode(path);
            }
            else
            {
                helper.SetXmlNodeString(path, s);
            }
        }
        /// <summary>
        /// Sets the value
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The x path</param>
        /// <param name="v">The boolean value</param>
        internal protected void SetValueBool(XmlHelper helper, string path, bool? v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                helper.SetXmlNodeBool(path, (bool)v);
            }
        }
        protected internal string GetAsString(object v)
        {
            return (v ?? "").ToString();
        }
        /// <summary>
        /// Is this value allowed to be changed?
        /// </summary>
        protected internal bool AllowChange { get; set; }

        internal ExcelDxfColor GetColor(XmlHelper helper, string path, eStyleClass styleClass)
        {
            ExcelDxfColor ret = new ExcelDxfColor(_styles, styleClass, _callback);
            ret.Theme = (eThemeSchemeColor?)helper.GetXmlNodeIntNull(path + "/@theme");
            ret.Index = helper.GetXmlNodeIntNull(path + "/@indexed");
            string rgb = helper.GetXmlNodeString(path + "/@rgb");
            if (rgb != "")
            {
                ret.Color = Color.FromArgb(int.Parse(rgb.Replace("#", ""), NumberStyles.HexNumber));
            }
            ret.Auto = helper.GetXmlNodeBoolNullable(path + "/@auto");
            ret.Tint = helper.GetXmlNodeDoubleNull(path + "/@tint");
            return ret;
        }
        internal static ExcelUnderLineType? GetUnderLineEnum(string value)
        {
            switch (value.ToLower(CultureInfo.InvariantCulture))
            {
                case "single":
                    return ExcelUnderLineType.Single;
                case "double":
                    return ExcelUnderLineType.Double;
                case "singleaccounting":
                    return ExcelUnderLineType.SingleAccounting;
                case "doubleaccounting":
                    return ExcelUnderLineType.DoubleAccounting;
                default:
                    return null;
            }
        }
    }
}
