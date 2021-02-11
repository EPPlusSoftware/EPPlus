/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table area of selection used for different purposes.
    /// </summary>
    public class ExcelPivotArea : XmlHelper
    {
        internal ExcelPivotArea(XmlNamespaceManager nsm, XmlNode topNode) : 
            base(nsm, topNode)
        {

        }
        /// <summary>
        /// The field referenced. A negative value is no field
        /// </summary>
        internal int FieldIndex
        { 
            get
            {
                return GetXmlNodeInt("@field");
            }
            set
            {
                SetXmlNodeInt("@field", value);
            }
        }
        /// <summary>
        /// Position of the field within the axis to which this rule applies. A negative value is no field position
        /// </summary>
        internal int FieldPosition 
        {
            get
            {
                return GetXmlNodeInt("@fieldPosition");
            }
            set
            {
                SetXmlNodeInt("@fieldPosition", value);
            }
        }

        /// <summary>
        /// The pivot area effected.
        /// </summary>
        public ePivotAreaType PivotAreaType
        {
            get
            {
                return GetXmlNodeString("@type").ToPivotAreaType();
            }
            set
            {
                if(value==ePivotAreaType.Normal)
                {
                    ((XmlElement)TopNode).RemoveAttribute("@type");
                }
                else
                {
                    SetXmlNodeString("@type", value.ToPivotAreaTypeString());
                }
            }
        }
        /// <summary>
        /// The region of the PivotTable effected.
        /// </summary>
        public ePivotTableAxis Axis 
        { 
            get
            {
                return GetXmlNodeString("@axis").ToPivotTableAxis();
            }
            set
            {
                SetXmlNodeString("@axis", value.ToPivotTableAxisString(), true);
            }
        }

        /// <summary>
        /// If the data values in the data area are included.
        /// <seealso cref="LabelOnly"/>
        /// </summary>
        public bool DataOnly 
        { 
            get
            {
                return GetXmlNodeBool("@dataOnly", true);
            }
            set
            {
                SetXmlNodeBool("@dataOnly", value);
            }
        }
        /// <summary>
        /// If the item labels are included. 
        /// <seealso cref="DataOnly"/>
        /// </summary>
        public bool LabelOnly
        {
            get
            {
                return GetXmlNodeBool("@labelOnly");
            }
            set
            {
                if(value && (PivotAreaType==ePivotAreaType.Data || PivotAreaType == ePivotAreaType.Normal /*|| PivotAreaType==ePivotAreaType.Origin */|| PivotAreaType==ePivotAreaType.TopEnd))
                {
                    throw (new InvalidOperationException());
                }
                SetXmlNodeBool("@labelOnly", value);
            }
        }
        /// <summary>
        /// If the row grand total is included
        /// </summary>
        public bool GrandRow
        {
            get
            {
                return GetXmlNodeBool("@grandRow");
            }
            set
            {
                SetXmlNodeBool("@grandRow", value);
            }
        }
        /// <summary>
        /// If the column grand total is included
        /// </summary>
        public bool GrandColumn
        {
            get
            {
                return GetXmlNodeBool("@grandCol");
            }
            set
            {
                SetXmlNodeBool("@grandCol", value);
            }
        }
        /// <summary>
        /// If any indexes refer to fields or items in the Pivot cache and not the view.
        /// </summary>
        public bool CacheIndex
        {
            get
            {
                return GetXmlNodeBool("@cacheIndex", true);
            }
            set
            {
                SetXmlNodeBool("@cacheIndex", value, true);
            }
        }
        /// <summary>
        /// Indicating whether the pivot table area referes to an area that is in outline mode.
        /// </summary>
        public bool Outline
        {
            get
            {
                return GetXmlNodeBool("@outline", true);
            }
            set
            {
                SetXmlNodeBool("@outline", value, true);
            }
        }
        /// <summary>
        /// A Reference that specifies a subset of the selection area. Points are relative to the top left of the selection area.
        /// </summary>
        public string Offset
        {
            get
            {
                return GetXmlNodeString("@offset");
            }
            internal set
            {
                SetXmlNodeString("@offset", value, true);
            }
        }
        /// <summary>
        /// If collapsed levels/dimensions are considered subtotals
        /// </summary>
        public bool CollapsedLevelsAreSubtotals 
        {
            get
            {
                return GetXmlNodeBool("@collapsedLevelsAreSubtotals");
            }
            set
            {
                SetXmlNodeBool("@collapsedLevelsAreSubtotals", value, false);
            }
        }
    }
}
