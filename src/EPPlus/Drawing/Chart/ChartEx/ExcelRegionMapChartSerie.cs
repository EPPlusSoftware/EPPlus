/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB            EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A chart series for a region map chart
    /// </summary>
    public class ExcelRegionMapChartSerie : ExcelChartExSerie
    {
        internal ExcelRegionMapChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {

        }

        const string _attributionPath = "cx:layoutPr/cx:geography/@attribution";
        /// <summary>
        /// The provider or source of the geographical data. Default is Bing.
        /// </summary>
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
        /// <summary>
        /// Specifies the country code. Uses the TwoLetterISOLanguageName property of the CultureInfo object.
        /// </summary>
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
        /// <summary>
        /// Specifies the language. 
        /// </summary>
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
        /// <summary>
        /// The cartographic map projection for the series
        /// </summary>
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
        /// <summary>
        /// The level of view for the series
        /// </summary>
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
        /// <summary>
        /// Colors for the gradient scale of the region map series. 
        /// </summary>
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
        /// <summary>
        /// Layout type for region labels
        /// </summary>
        public eRegionLabelLayout RegionLableLayout
        {
            get
            {
                return GetXmlNodeString("cx:layoutPr/cx:regionLabelLayout/@val").ToEnum(eRegionLabelLayout.None);
            }
            set
            {
                SetXmlNodeString("cx:layoutPr/cx:regionLabelLayout/@val", value.ToEnumString());
            }
        }

        public eColorBy ColorBy 
        { 
            get
            {
                if(DataDimensions.Count < 2 && DataDimensions[0] is ExcelChartExStringData s)
                {
                    if(s.Type==eStringDataType.ColorString)
                    {
                        return eColorBy.CategoryNames;
                    }
                }
                return eColorBy.Value;
            }
            set
            {
                if(ColorBy != value)
                {
                    if(value==eColorBy.Value)
                    {
                        DataDimensions.SetTypeNumeric(0, eNumericDataType.ColorValue);
                    }
                    else
                    {
                        DataDimensions.SetTypeString(0, eStringDataType.ColorString);
                    }
                }
            }
        }
    }
}
