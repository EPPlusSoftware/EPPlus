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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableFieldItem
    {
        public ExcelPivotTableFieldItem()
        {
        }
        public ExcelPivotTableFieldItem (XmlElement node)
        {
            foreach(XmlAttribute a in node.Attributes)
            {
                switch(a.LocalName)
                {
                    case "c":
                        C = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "d":
                        D = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "e":
                        E = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "f":
                        F = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "h":
                        Hidden = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "m":
                        M = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "n":
                        Text = a.Value;
                        break;
                    case "s":
                        S = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "sd":
                        HideDetails = XmlHelper.GetBoolFromString(a.Value);
                        break;
                    case "t":                        
                        Type = a.Value.ToEnum(eItemType.Data);
                        break;
                    case "x":
                        X = int.Parse(a.Value);
                        break;
                }
            }
        }
        public string Text { get; set; }
        public object Value { get; internal set; }
        public bool Hidden { get; set; }
        internal bool HideDetails { get; set; } = true;
        internal bool C { get; set; }
        internal bool D { get; set; }
        internal bool E { get; set; } = true;
        internal bool F { get; set; }
        internal bool M { get; set; }
        internal bool S { get; set; }
        internal int X { get; set; } = -1;
        internal eItemType Type { get; set; }

        internal void GetXmlString(StringBuilder sb)
        {
            sb.Append("<item");
            if(X>-1)
            {
                sb.AppendFormat(" x={0}", X);
            }
            if(Type!=eItemType.Data)
            {
                sb.AppendFormat(" T={0}", Type.ToEnumString());
            }
            if(!string.IsNullOrEmpty(Text))
            {
                sb.AppendFormat(" x=\"{0}\"", Text);
            }
            AddBool(sb,"h", Hidden);
            AddBool(sb, "sd", HideDetails, true);
            AddBool(sb, "c", C);
            AddBool(sb, "d", D);
            AddBool(sb, "e", E, true);
            AddBool(sb, "f", F);
            AddBool(sb, "m", M);
            AddBool(sb, "s", S);
            sb.Append("/>");
        }

        private void AddBool(StringBuilder sb, string attrName, bool b, bool defaultValue=false)
        {
            if(b != defaultValue)
            {
                sb.AppendFormat(" {0}={1}",attrName, b?"1":"0");
            }
        }
    }
    /// <summary>
    /// A field Item. Used for grouping
    /// </summary>
    //public class ExcelPivotTableFieldItem : XmlHelper
    //{
    //ExcelPivotTableField _field;
    //internal ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
    //    base(ns, topNode)
    //{
    //   _field = field;
    //}
    ///// <summary>
    ///// The text. Unique values only
    ///// </summary>
    //public string Text
    //{
    //    get
    //    {
    //        return GetXmlNodeString("@n");
    //    }
    //    set
    //    {
    //        if(string.IsNullOrEmpty(value))
    //        {
    //            DeleteNode("@n");
    //            return;
    //        }
    //        foreach (var item in _field.Items)
    //        {
    //            if (item.Text == value)
    //            {
    //                throw(new ArgumentException("Duplicate Text"));
    //            }
    //        }
    //        SetXmlNodeString("@n", value);
    //    }
    //}
    //internal int X
    //{
    //    get
    //    {
    //        return GetXmlNodeInt("@x"); 
    //    }
    //}
    //internal string T
    //{
    //    get
    //    {
    //        return GetXmlNodeString("@t");
    //    }
    //}
    //public bool HideDetails 
    //{
    //    get
    //    {
    //        return GetXmlNodeBool("@sd");
    //    }
    //    set
    //    {
    //        SetXmlNodeBool("@sd", value, false);
    //    }
    //}
    //internal bool Hidden
    //{
    //    get
    //    {
    //        return GetXmlNodeBool("@h");
    //    }
    //    set
    //    {
    //        SetXmlNodeBool("@h", value, false);
    //    }
    //}
    //}
}
