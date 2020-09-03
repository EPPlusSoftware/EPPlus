/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.IO;
using System.Xml;

namespace EPPlusTest.Table.PivotTable.Filter
{
    public class ExcelPivotTableFilter : XmlHelper
    {
        internal ExcelPivotTableFilter(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            if(topNode.InnerXml=="")
            {
                topNode.InnerXml= "<autofilter ref=\"A1\"><filterColumn col=\"0\"><filters/></filterColumn</autofilter>";
            }
        }
        public int Id
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            internal set
            {
                SetXmlNodeInt("@id", value);
            }
        }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value, true);
            }
        }
        public string Description
        {
            get
            {
                return GetXmlNodeString("@description");
            }
            set
            {
                SetXmlNodeString("@description", value, true);
            }
        }
        public ePivotTableFilterType Type
        {
            get
            {
                return GetXmlNodeString("@type").ToEnum(ePivotTableFilterType.Unknown);
            }
            internal set
            {
                SetXmlNodeString("@type", value.ToEnumString());
            }
        }
        public int EvalOrder
        {
            get
            {
                return GetXmlNodeInt("@evalOrder");
            }
            internal set
            {
                SetXmlNodeInt("@evalOrder", value);
            }
        }
        internal int Fld
        {
            get
            {
                return GetXmlNodeInt("@fld");
            }
            set
            {
                SetXmlNodeInt("@fld", value);
            }
        }
        internal int MeasureFldIndex
        {
            get
            {
                return GetXmlNodeInt("@iMeasureFld");
            }
            set
            {
                SetXmlNodeInt("@iMeasureFld", value);
            }
        }
        internal int MeasureHierIndex
        {
            get
            {
                return GetXmlNodeInt("@iMeasureHier");
            }
            set
            {
                SetXmlNodeInt("@iMeasureHier", value);
            }
        }
        internal int MemberPropertyFldIndex
        {
            get
            {
                return GetXmlNodeInt("@mpFld");
            }
            set
            {
                SetXmlNodeInt("@mpFld", value);
            }
        }
        
        internal string StringValue1
        {
            get
            {
                return GetXmlNodeString("@stringValue1");
            }
            set
            {
                SetXmlNodeString("@stringValue1", value, true);
            }
        }
        internal string StringValue2
        {
            get
            {
                return GetXmlNodeString("@stringValue2");
            }
            set
            {
                SetXmlNodeString("@stringValue2", value, true);
            }
        }
        ExcelFilterColumn _filter = null;
        public ExcelFilterColumn Filter
        {
            get
            {
                if(_filter==null)
                {
                    var filterNode = GetNode("d:autoFilter/d:filterColumn");
                    if(filterNode!=null)
                    {
                        switch(filterNode.LocalName)
                        {
                            case "customFilters":
                                _filter = new ExcelCustomFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "top10":
                                _filter = new ExcelTop10FilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "filters":
                                _filter = new ExcelValueFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "dynamicFilter":
                                _filter = new ExcelDynamicFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "colorFilter":
                                _filter = new ExcelColorFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "iconFilter":
                                _filter = new ExcelIconFilterColumn(NameSpaceManager, filterNode);
                                break;
                            default:
                                _filter = null;
                                break;
                        }
                    }
                    else
                    {
                        throw new Exception("Invalid xml in pivot table. Missing Filter column");
                    }
                }
                return _filter;
            }
        }
    }
}
