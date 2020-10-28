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
    public class ExcelControlScrollBar : ExcelControl
    {
        internal ExcelControlScrollBar(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, rel,  controlPropertiesXml, null)
        {
        }

        public override eControlType ControlType => eControlType.ScrollBar;
        /// <summary>
        /// Gets or sets the address to the cell that is linked to the control. 
        /// </summary>
        public ExcelAddressBase LinkedCell
        {
            get
            {
                return LinkedCellBase;
            }
            set
            {
                LinkedCellBase = value;
            }
        }
        /// <summary>
        /// Gets or sets if scrollbar is horizontal or vertical
        /// </summary>
        public bool Horizontal
        {
            get
            {
                return _ctrlProp.GetXmlNodeBool("@horiz");
            }
            set
            {
                _ctrlProp.SetXmlNodeBool("@horiz", value);
                _vmlProp.SetXmlNodeBool("x:Horiz", value);
            }
        }
        /// <summary>
        /// How much the scrollbar is incremented for each click
        /// </summary>
        public int Increment
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@inc", 1);
            }
            set
            {
                if(value < 0 || value >3000)
                {
                    throw (new ArgumentOutOfRangeException("Increment must be between 0 and 3000"));
                }
                _ctrlProp.SetXmlNodeInt("@inc", value);
                _vmlProp.SetXmlNodeInt("x:Inc", value);
            }
        }
        /// <summary>
        /// The number of items to move the scroll bar on a page click. Null is default
        /// </summary>
        public int? Page
        {
            get
            {
                return _ctrlProp.GetXmlNodeIntNull("@page");
            }
            set
            {
                if (value.HasValue && (value < 0 || value > 3000))
                {
                    throw (new ArgumentOutOfRangeException("Page must be between 0 and 3000"));
                }
                _ctrlProp.SetXmlNodeInt("@page", value);
                _vmlProp.SetXmlNodeInt("x:Page", value);
            }
        }

    }
}