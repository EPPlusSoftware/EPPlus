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
using OfficeOpenXml.Packaging;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public class ExcelControlEditBox : ExcelControlWithText
    {
        internal ExcelControlEditBox(ExcelDrawings drawings, XmlElement drawNode) : base(drawings, drawNode)
        {
        }
        internal ExcelControlEditBox(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, part, controlPropertiesXml, null)
        {
        }

        public override eControlType ControlType => eControlType.EditBox;
        /// <summary>
        /// The source data cell that the control object's data is linked to.
        /// </summary>
        public ExcelAddressBase LinkedCell
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaTxbx");
                if (ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaTxbx");
                    _vmlProp.DeleteNode("x:FmlaTxbx");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaTxbx", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaTxbx", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaTxbx", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaTxbx", value.FullAddress);
                }
            }
        }

        //editVal
        //multiLine
        //verticalBar
        //passwordEdit
    }
}