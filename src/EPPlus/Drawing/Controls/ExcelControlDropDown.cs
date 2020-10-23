/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Packaging;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public class ExcelControlDropDown : ExcelControlList
    {
        internal ExcelControlDropDown(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, rel,  controlPropertiesXml)
        {
        }

        public override eControlType ControlType => eControlType.DropDown;
        /// <summary>
        /// Get or sets whether a drop-down object has a color applied to it
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
                _ctrlProp.SetXmlNodeInt("@dropLines", value);
            }
        }
        /// <summary>
        /// The style of the drop-down.
        /// </summary>
        public eDropStyle DropStyle
        {
            get
            {
                switch(GetXmlNodeString("@dropStyle"))
                {
                    case "comboedit":
                        return eDropStyle.Combo;
                    case "simple":
                        return eDropStyle.Simple;
                    default:
                        return eDropStyle.ComboEdit;
                }
            }
            set
            {
                SetXmlNodeString("@dropStyle", value.ToString().ToLower());
            }
        }
    }
}