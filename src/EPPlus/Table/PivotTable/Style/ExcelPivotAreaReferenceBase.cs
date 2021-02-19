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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public abstract class ExcelPivotAreaReferenceBase : XmlHelper
    {
        internal protected ExcelPivotTable _pt;
        internal ExcelPivotAreaReferenceBase(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) : base(nsm, topNode)
        {
            _pt = pt;
        }
        internal int FieldIndex
        { 
            get
            {
                var v=GetXmlNodeLong("@field");
                if(v > int.MaxValue)
                {
                    return -2;
                }
                else
                {
                    return (int)v;
                }
            }
            set
            {
                if(value<0)
                {
                    SetXmlNodeLong("@field", 4294967294);
                }
                else
                {
                    SetXmlNodeInt("@field", value);
                }
            }
        }
        public bool Selected 
        {
            get
            {
                return GetXmlNodeBool("@selected", true);
            }
            set
            {
                SetXmlNodeBool("@selected", value);
            }
        }
        internal bool Relative 
        { 
            get
            {
                return GetXmlNodeBool("@relative");
            }
            set
            {
                SetXmlNodeBool("@relative", value);
            }
        }
        internal bool ByPosition 
        {
            get
            {
                return GetXmlNodeBool("@byPosition");
            }
            set
            {
                SetXmlNodeBool("@byPosition", value);
            }
        }
        internal abstract void UpdateXml();
        public bool DefaultSubtotal 
        { 
            get
            {
                return GetXmlNodeBool("@defaultSubtotal");
            }
            set
            {
                SetXmlNodeBool("@defaultSubtotal", value);
            }
        }
        public bool AvgSubtotal
        {
            get
            {
                return GetXmlNodeBool("@avgSubtotal");
            }
            set
            {
                SetXmlNodeBool("@avgSubtotal", value);
            }
        }
        public bool CountSubtotal
        {
            get
            {
                return GetXmlNodeBool("@countSubtotal");
            }
            set
            {
                SetXmlNodeBool("@countSubtotal", value);
            }
        }
        public bool CountASubtotal
        {
            get
            {
                return GetXmlNodeBool("@countASubtotal");
            }
            set
            {
                SetXmlNodeBool("@countASubtotal", value);
            }
        }
        public bool MaxSubtotal
        {
            get
            {
                return GetXmlNodeBool("@maxSubtotal");
            }
            set
            {
                SetXmlNodeBool("@maxSubtotal", value);
            }
        }
        public bool MinSubtotal
        {
            get
            {
                return GetXmlNodeBool("@minSubtotal");
            }
            set
            {
                SetXmlNodeBool("@minSubtotal", value);
            }
        }
        public bool ProductSubtotal
        {
            get
            {
                return GetXmlNodeBool("@productSubtotal");
            }
            set
            {
                SetXmlNodeBool("@productSubtotal", value);
            }
        }
        public bool StdDevPSubtotal
        {
            get
            {
                return GetXmlNodeBool("@StdDevPSubtotal");
            }
            set
            {
                SetXmlNodeBool("@StdDevPSubtotal", value);
            }
        }
        public bool StdDevSubtotal
        {
            get
            {
                return GetXmlNodeBool("@StdDevSubtotal");
            }
            set
            {
                SetXmlNodeBool("@StdDevSubtotal", value);
            }
        }
        public bool SumSubtotal
        {
            get
            {
                return GetXmlNodeBool("@sumSubtotal");
            }
            set
            {
                SetXmlNodeBool("@sumSubtotal", value);
            }
        }
        public bool VarPSubtotal
        {
            get
            {
                return GetXmlNodeBool("@varPSubtotal");
            }
            set
            {
                SetXmlNodeBool("@varPSubtotal", value);
            }
        }
        public bool VarSubtotal
        {
            get
            {
                return GetXmlNodeBool("@varSubtotal");
            }
            set
            {
                SetXmlNodeBool("@varSubtotal", value);
            }
        }
        internal void SetFunction(DataFieldFunctions function)
        {
            switch(function)
            {
                case DataFieldFunctions.Average:
                    AvgSubtotal = true;
                    break;
                case DataFieldFunctions.Count:
                    CountSubtotal = true;
                    break;
                case DataFieldFunctions.CountNums:
                    CountASubtotal = true;
                    break;
                case DataFieldFunctions.Max:
                    MaxSubtotal = true;
                    break;
                case DataFieldFunctions.Min:
                    MinSubtotal = true;
                    break;
                case DataFieldFunctions.Product:
                    ProductSubtotal = true;
                    break;
                case DataFieldFunctions.StdDevP:
                    StdDevPSubtotal = true;
                    break;
                case DataFieldFunctions.StdDev:
                    StdDevSubtotal = true;
                    break;
                case DataFieldFunctions.Sum:
                    SumSubtotal = true;
                    break;
                case DataFieldFunctions.VarP:
                    VarPSubtotal = true;
                    break;
                case DataFieldFunctions.Var:
                    VarSubtotal = true;
                    break;
                default:
                    DefaultSubtotal = true;
                    break;
            }
        }
    }
}