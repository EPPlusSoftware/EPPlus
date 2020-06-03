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
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Geomapping level
    /// </summary>
    public enum eGeoMappingLevel
    {
        /// <summary>
        /// Geomapping level is handled automatic
        /// </summary>
        Automatic,
        /// <summary>
        /// Only regions which correspond to data points in the geographical category of a geospatial series are in view.
        /// </summary>
        DataOnly,
        /// <summary>
        /// The level of view for the series is set to postal code.
        /// </summary>
        PostalCode,
        /// <summary>
        /// The level of view for the series is set to county.
        /// </summary>
        County,
        /// <summary>
        /// The level of view for the series is set to state or province.
        /// </summary>
        State,
        /// <summary>
        /// The level of view for series is set to country/region.
        /// </summary>
        CountryRegion,
        /// <summary>
        /// The level of view for the series is set to continent.
        /// </summary>
        CountryRegionList,
        /// <summary>
        /// The level of view for the series is set to the entire world.
        /// </summary>
        World
    }
}