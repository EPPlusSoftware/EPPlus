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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Core.CellStore;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartExSerie : ExcelChartSerieBase
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
       internal ExcelChartExSerie(ExcelChartBase chart, XmlNamespaceManager ns, XmlNode node)
           : base(chart,ns,node)
       {
       }
       const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
       /// Header address for the serie.
       /// </summary>
       public override ExcelAddressBase HeaderAddress
       {
           get
           {
              return null; //TODO check handling
           }
           set
           {
                throw new NotImplementedException();
            }
       }        
        /// <summary>
        /// Set this to a valid address or the drawing will be invalid.
        /// </summary>
        public override string Series
        {
           get
           {
               return GetXmlNodeString("cx:*Dim[type='val']");
           }
           set
           {
                SetXmlNodeString("cx:*Dim[type='val']", value); 
           }

       }
        /// <summary>
        /// Set an address for the horisontal labels
        /// </summary>
        public override string XSeries
       {
           get
           {
                return GetXmlNodeString("cx:*Dim[type='cat']");
            }
           set
           {
                SetXmlNodeString("cx:*Dim[type='cat']", value);
            }
       }

        public override string Header { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
    }
}
