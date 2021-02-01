﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Represents a drop down form control
    /// </summary>
    public class ExcelControlDropDown : ExcelControlList
    {
        internal ExcelControlDropDown(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
            SetSize(150, 20); //Default size
        }
        internal ExcelControlDropDown(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }

        /// <summary>
        /// The type of form control
        /// </summary>
        public override eControlType ControlType => eControlType.DropDown;
        /// <summary>
        /// Gets or sets whether a drop-down object has a color applied to it
        /// </summary>
        public bool Colored 
        {
            get
            {
                return _ctrlProp.GetXmlNodeBool("@colored");
            }
            set
            {
                _ctrlProp.SetXmlNodeBool("@colored", value);
                _vmlProp.SetXmlNodeBool("x:Colored", value);
            }
        }
        /// <summary>
        /// Gets or sets the number of lines before a scroll bar is added to the drop-down.
        /// </summary>
        public int DropLines
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@dropLines", 8);
            }
            set
            {
                _ctrlProp.SetXmlNodeInt("@dropLines", value, null, false);
                _vmlProp.SetXmlNodeInt("x:DropLines", value);
            }
        }
        /// <summary>
        /// The style of the drop-down.
        /// </summary>
        public eDropStyle DropStyle
        {
            get
            {
                switch(_ctrlProp.GetXmlNodeString("@dropStyle"))
                {
                    case "comboedit":
                        return eDropStyle.ComboEdit;
                    case "simple":
                        return eDropStyle.Simple;
                    default:
                        return eDropStyle.Combo;
                }
            }
            set
            {
                _ctrlProp.SetXmlNodeString("@dropStyle", value.ToString().ToLower());
                _vmlProp.SetXmlNodeString("x:DropStyle", value.ToString());
            }
        }
        /// <summary>
        /// Minimum width 
        /// </summary>
        public int? MinimumWidth
        {
            get
            {
                return _ctrlProp.GetXmlNodeIntNull("@widthMin");
            }
            set
            {
                _ctrlProp.SetXmlNodeInt("@widthMin", value,null, false);
                _ctrlProp.SetXmlNodeInt("x:WidthMin", value);
            }
        }
    }
}