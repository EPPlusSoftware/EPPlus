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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot table field Item. Used for grouping.
    /// </summary>
    public class ExcelPivotTableFieldItem
    {
        [Flags]
        internal enum eBoolFlags
        {
            Hidden=0x1,
            ShowDetails = 0x2,
            C = 0x4,
            D = 0x8,
            E = 0x10,
            F = 0x20,
            M = 0x40,
            S = 0x80
        }
        internal eBoolFlags flags=eBoolFlags.ShowDetails|eBoolFlags.E;
        internal ExcelPivotTableFieldItem()
        {
        }
        internal ExcelPivotTableFieldItem (XmlElement node)
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
                        ShowDetails = XmlHelper.GetBoolFromString(a.Value);
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
        /// <summary>
        /// The custom text of the item. Unique values only
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// The value of the item
        /// </summary>
        public object Value { get; internal set; }
        /// <summary>
        /// A flag indicating if the items are hidden
        /// </summary>
        public bool Hidden 
        { 
            get
            {
                return (flags & eBoolFlags.Hidden) == eBoolFlags.Hidden;
            }
            set
            {
                if (Type != eItemType.Data) throw (new InvalidOperationException("Hidden can only be set for items of type Data"));
                SetFlag(eBoolFlags.Hidden, value);
            }
        }

        /// <summary>
        /// A flag indicating if the items expanded or collapsed.
        /// </summary>
        public bool ShowDetails
        {
            get
            {
                return (flags & eBoolFlags.ShowDetails) == eBoolFlags.ShowDetails;
            }
            set
            {
                SetFlag(eBoolFlags.ShowDetails, value);
            }
        }
        internal bool C
        {
            get
            {
                return (flags & eBoolFlags.C) == eBoolFlags.C;
            }
            set
            {
                SetFlag(eBoolFlags.C, value);
            }
        }
        internal bool D
        {
            get
            {
                return (flags & eBoolFlags.D) == eBoolFlags.D;
            }
            set
            {
                SetFlag(eBoolFlags.D, value);
            }
        }
        internal bool E
        {
            get
            {
                return (flags & eBoolFlags.E) == eBoolFlags.E;
            }
            set
            {
                SetFlag(eBoolFlags.E, value);
            }
        }
        internal bool F
        {
            get
            {
                return (flags & eBoolFlags.F) == eBoolFlags.F;
            }
            set
            {
                SetFlag(eBoolFlags.F, value);
            }
        }
        internal bool M
        {
            get
            {
                return (flags & eBoolFlags.M) == eBoolFlags.M;
            }
            set
            {
                SetFlag(eBoolFlags.M, value);
            }
        }
        internal bool S
        {
            get
            {
                return (flags & eBoolFlags.S) == eBoolFlags.S;
            }
            set
            {
                SetFlag(eBoolFlags.S, value);
            }
        }
        internal int X { get; set; } = -1;
        internal eItemType Type { get; set; }

        internal void GetXmlString(StringBuilder sb)
        {
            if (X == -1 && Type == eItemType.Data) return;
            sb.Append("<item");
            if(X>-1)
            {
                sb.AppendFormat(" x=\"{0}\"", X);
            }
            if(Type!=eItemType.Data)
            {
                sb.AppendFormat(" t=\"{0}\"", Type.ToEnumString());
            }
            if(!string.IsNullOrEmpty(Text))
            {
                sb.AppendFormat(" n=\"{0}\"", OfficeOpenXml.Utils.ConvertUtil.ExcelEscapeString(Text));
            }
            AddBool(sb,"h", Hidden);
            AddBool(sb, "sd", ShowDetails, true);
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
                sb.AppendFormat(" {0}=\"{1}\"",attrName, b?"1":"0");
            }
        }
        private void SetFlag(eBoolFlags flag, bool value)
        {
            if(value)
            {
                flags |= flag;
            }
            else
            {
                flags &= ~flag;
            }
        }
    }
}
