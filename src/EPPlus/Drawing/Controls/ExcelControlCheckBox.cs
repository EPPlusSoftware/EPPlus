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
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Represents a check box form control
    /// </summary>
    public class ExcelControlCheckBox : ExcelControlWithColorsAndLines
    {
        internal ExcelControlCheckBox(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
        }

        internal ExcelControlCheckBox(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }

        /// <summary>
        /// The type of form control
        /// </summary>
        public override eControlType ControlType => eControlType.CheckBox;
        /// <summary>
        /// Gets or sets the state of a check box 
        /// </summary>
        public eCheckState Checked 
        { 
            get
            {
                return _ctrlProp.GetXmlNodeString("@checked").ToEnum(eCheckState.Unchecked);
            }
            set
            {
                if(value==eCheckState.Unchecked)
                {
                    _ctrlProp.DeleteNode("@checked");
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@checked", value.ToString());
                }
                _vmlProp.SetXmlNodeInt("x:Checked",(int)value);
                if(LinkedCell!=null)
                {
                    ExcelWorksheet ws;
                    if(string.IsNullOrEmpty(LinkedCell.WorkSheetName))
                    {
                        ws = _drawings.Worksheet;
                    }
                    else
                    {
                        ws = _drawings.Worksheet.Workbook.Worksheets[LinkedCell.WorkSheetName];
                    }

                    if (ws!=null)
                    {
                        if(value == eCheckState.Checked)
                        {
                            ws.Cells[LinkedCell.Address].Value = true;
                        }
                        else if (value == eCheckState.Unchecked)
                        {
                            ws.Cells[LinkedCell.Address].Value = false;
                        }
                        else
                        {
                            ws.Cells[LinkedCell.Address].Value = ExcelErrorValue.Create(eErrorType.NA);
                        }                           
                    }
                }
            }
        }
    }
}