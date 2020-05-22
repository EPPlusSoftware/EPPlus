/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelRegionMapChartSerie : ExcelChartExSerie
    {
        internal ExcelRegionMapChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {

        }

        const string _attributionPath = "cx:layoutPr/cx:geography/@attribution";
        public string Attribution 
        { 
            get
            {
                return GetXmlNodeString(_attributionPath);
            }
            set
            {
                SetXmlNodeString(_attributionPath, value);
            }
        }
        const string _regionPath = "cx:layoutPr/cx:geography/@cultureRegion";
        public CultureInfo Region 
        { 
            get
            {
                var r=GetXmlNodeString(_regionPath);
                return new CultureInfo(r);
            }
            set
            {
                if(value==null || value.TwoLetterISOLanguageName.Length!=2)
                {
                    throw (new InvalidOperationException("Region must have a two letter ISO code"));
                }
                SetXmlNodeString(_regionPath, value.TwoLetterISOLanguageName);
            }
        }

        const string _languagePath = "cx:layoutPr/cx:geography/@cultureLanguage";
        public CultureInfo Language 
        {
            get
            {
                var r = GetXmlNodeString(_languagePath);
                return new CultureInfo(r);
            }
            set
            {
                if (value == null)
                {
                    throw (new InvalidOperationException("Language must not be null."));
                }
                SetXmlNodeString(_languagePath, value.Name);
            }
        }
        const string _projectionTypePath = "cx:layoutPr/cx:geography/@projectionType";
        public eProjectionType ProjectionType 
        { 
            get
            {
                return GetXmlNodeString(_projectionTypePath).ToEnum(eProjectionType.Automatic);
            }
            set
            {
                if (value == eProjectionType.Automatic)
                {
                    DeleteNode(_projectionTypePath);
                }
                else
                {
                    SetXmlNodeString(_projectionTypePath, value.ToEnumString());
                }
            }
        }
        const string _geoMappingLevelPath = "cx:layoutPr/cx:geography/@viewedRegionType";
        public eGeoMappingLevel ViewedRegionType
        {
            get
            {
                return GetXmlNodeString(_geoMappingLevelPath).ToEnum(eGeoMappingLevel.Automatic);
            }
            set
            {
                if(value==eGeoMappingLevel.Automatic)
                {
                    DeleteNode(_geoMappingLevelPath);
                }
                else
                {
                    SetXmlNodeString(_geoMappingLevelPath, value.ToEnumString());
                }                
            }
        }
        ExcelChartExValueColors _colors = null;
        public ExcelChartExValueColors Colors
        {
            get
            {
                if(_colors==null)
                {
                    _colors = new ExcelChartExValueColors(this, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _colors;
            }
        }
    }
}
