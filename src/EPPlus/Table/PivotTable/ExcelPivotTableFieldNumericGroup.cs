/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot table field numeric grouping
    /// </summary>
    public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
    {
        internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
		}
        const string startPath = "d:rangePr/@startNum";
        /// <summary>
        /// Start value
        /// </summary>
        public double Start
        {
            get
            {
                return GetXmlNodeDoubleNull(startPath) ?? 0D;
            }
            private set
            {
                SetXmlNodeString(startPath,value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string endPath = "d:rangePr/@endNum";
        /// <summary>
        /// End value
        /// </summary>
        public double End
        {
            get
            {
                return GetXmlNodeDoubleNull(endPath)??0;
            }
            private set
            {
                SetXmlNodeString(endPath, value.ToString(CultureInfo.InvariantCulture));
                CalculateEndIsDivisibleWithInterval();
			}
        }
        const string groupIntervalPath = "d:rangePr/@groupInterval";
        /// <summary>
        /// Interval
        /// </summary>
        public double Interval
        {
            get
            {
                return GetXmlNodeDoubleNull(groupIntervalPath) ?? 0D;
            }
            private set
			{
				SetXmlNodeString(groupIntervalPath, value.ToString(CultureInfo.InvariantCulture));
				CalculateEndIsDivisibleWithInterval();
			}
		}

		internal void CalculateEndIsDivisibleWithInterval()
		{
			EndIsDivisibleWithInterval = End % Interval == 0;
		}

		internal bool EndIsDivisibleWithInterval { get; private set; }
	}
}