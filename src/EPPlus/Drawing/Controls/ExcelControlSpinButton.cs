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
    /// <summary>
    /// Represents a spin button form control
    /// </summary>
    public class ExcelControlSpinButton : ExcelControl
    {
        internal ExcelControlSpinButton(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
            SetSize(40, 80); //Default size
        }

        internal ExcelControlSpinButton(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }

        /// <summary>
        /// The type of form control
        /// </summary>
        public override eControlType ControlType => eControlType.SpinButton;
        /// <summary>
        /// How much the spin button is incremented for each click
        /// </summary>
        public int Increment
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@inc", 1);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("Increment must be between 0 and 3000"));
                }
                _ctrlProp.SetXmlNodeInt("@inc", value);
                _vmlProp.SetXmlNodeInt("x:Inc", value);
            }
        }
        /// <summary>
        /// The value when a spin button is at it's minimum
        /// </summary>
        public int MinValue
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@min", 0);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("MinValue must be between 0 and 3000"));
                }
                _ctrlProp.SetXmlNodeInt("@min", value);
                _vmlProp.SetXmlNodeInt("x:Min", value);
            }
        }
        /// <summary>
        /// The value when a spin button is at it's maximum
        /// </summary>
        public int MaxValue
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@max", 30000);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("MaxValue must be between 0 and 30000"));
                }
                _ctrlProp.SetXmlNodeInt("@max", value);
                _vmlProp.SetXmlNodeInt("x:Max", value);
            }
        }
        /// <summary>
        /// The value when a spin button is at it's maximum
        /// </summary>
        public int Value
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@val", 0);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 30000"));
                }
                _ctrlProp.SetXmlNodeInt("@val", value);
                _vmlProp.SetXmlNodeInt("x:Val", value);

                SetLinkedCellValue(value);
            }
        }
    }
}