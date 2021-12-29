/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Core.CellStore;
using System.Globalization;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartStandardSerie : ExcelChartSerie
    {
        private readonly bool _isPivot;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
       /// <param name="isPivot">Is pivotchart</param>  
       internal ExcelChartStandardSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(chart, ns, node)
       {
           _chart = chart;
           _isPivot = isPivot;
           SchemaNodeOrder = new string[] { "idx", "order", "tx", "spPr", "marker", "invertIfNegative", "pictureOptions", "explosion", "dPt", "dLbls", "trendline","errBars", "cat", "val", "xVal", "yVal", "smooth","shape", "bubbleSize", "bubble3D", "numRef", "numLit", "strRef", "strLit", "formatCode", "ptCount", "pt" };

           if (_chart.ChartNode.LocalName=="scatterChart" ||
               _chart.ChartNode.LocalName.StartsWith("bubble", StringComparison.OrdinalIgnoreCase))
           {
               _seriesTopPath = "c:yVal";
               _xSeriesTopPath = "c:xVal";
           }
           else
           {
               _seriesTopPath = "c:val";
               _xSeriesTopPath = "c:cat";
           }

           _seriesPath = string.Format(_seriesPath, _seriesTopPath);
           _numCachePath = string.Format(_numCachePath, _seriesTopPath);

            var np = string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:numRef");
            var sp= string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:strRef");

            if(ExistsNode(sp))
            {
                _xSeriesPath = sp;
            }
            else
            {
                _xSeriesPath = np;
            }
            _seriesStrLitPath = string.Format("{0}/c:strLit", _seriesTopPath);
            _seriesNumLitPath = string.Format("{0}/c:numLit", _seriesTopPath);

            _xSeriesStrLitPath = string.Format("{0}/c:strLit", _xSeriesTopPath);
            _xSeriesNumLitPath = string.Format("{0}/c:numLit", _xSeriesTopPath);
       }       
       internal override void SetID(string id)
       {
           SetXmlNodeString("c:idx/@val",id);
           SetXmlNodeString("c:order/@val", id);
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public override string Header 
       {
           get
           {
                return GetXmlNodeString(headerPath);
            }
            set
            {
                Cleartx();
                SetXmlNodeString(headerPath, value);            
            }
        }

       private void Cleartx()
       {
           var n = TopNode.SelectSingleNode("c:tx", NameSpaceManager);
           if (n != null)
           {
               n.InnerXml = "";
           }
       }
       const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
       /// Header address for the serie.
       /// </summary>
       public override ExcelAddressBase HeaderAddress
       {
           get
           {
               string address = GetXmlNodeString(headerAddressPath);
               if (address == "")
               {
                   return null;
               }
               else
               {
                   return new ExcelAddressBase(address);
               }
            }
            set
            {
                if ((value._fromCol != value._toCol && value._fromRow != value._toRow) || value.Addresses != null) //Single cell removed, allow row & column --> issue 15102. 
                {
                    throw (new ArgumentException("Address must be a row, column or single cell"));
                }

                Cleartx();
                SetXmlNodeString(headerAddressPath, ExcelCellBase.GetFullAddress(value.WorkSheetName, value.Address));
                SetXmlNodeString("c:tx/c:strRef/c:strCache/c:ptCount/@val", "0");
            }
        }        
        string _seriesTopPath;
        string _seriesPath = "{0}/c:numRef/c:f";
        string _numCachePath = "{0}/c:numRef/c:numCache";
        string _seriesStrLitPath, _seriesNumLitPath;
        /// <summary>
        /// Set this to a valid address or the drawing will be invalid.
        /// </summary>
        public override string Series
        {
           get
           {
               return GetXmlNodeString(_seriesPath);
           }
           set
           {
                value = value.Trim();
                if (value.StartsWith("=", StringComparison.OrdinalIgnoreCase)) value = value.Substring(1);
                if (value.StartsWith("{", StringComparison.OrdinalIgnoreCase) && value.EndsWith("}", StringComparison.OrdinalIgnoreCase))
                {
                    GetLitValues(value, out double[] numLit, out string[] strLit);
                    if(strLit!=null)
                    {
                        throw (new ArgumentException("Value series can't contain strings"));
                    }
                    NumberLiteralsY = numLit;
                    SetLits(NumberLiteralsY, null, _seriesNumLitPath, _seriesStrLitPath);
                }
                else
                {
                    NumberLiteralsX = null;
                    StringLiteralsX = null;
                    SetSerieFunction(value);
                }
            }

        }

       string _xSeries=null;
       string _xSeriesTopPath;
       string _xSeriesPath = "{0}/{1}/c:f";
       string _xSeriesStrLitPath, _xSeriesNumLitPath;
        /// <summary>
        /// Set an address for the horisontal labels
        /// </summary>
       public override string XSeries
       {
           get
           {
               return GetXmlNodeString(_xSeriesPath);
           }
           set
           {
                _xSeries = value.Trim();
                if (_xSeries.StartsWith("=", StringComparison.OrdinalIgnoreCase)) _xSeries = _xSeries.Substring(1);
                if (value.StartsWith("{", StringComparison.OrdinalIgnoreCase) && value.EndsWith("}", StringComparison.OrdinalIgnoreCase))
                {
                    GetLitValues(_xSeries, out double[] numLit, out string[] strLit);
                    NumberLiteralsX = numLit;
                    StringLiteralsX = strLit;
                    SetLits(NumberLiteralsX, StringLiteralsX, _xSeriesNumLitPath, _xSeriesStrLitPath);
                }
                else
                {
                    NumberLiteralsX = null;
                    StringLiteralsX = null;
                    CreateNode(_xSeriesPath, true);
                    if(ExcelCellBase.IsValidAddress(_xSeries))
                    {
                        SetXmlNodeString(_xSeriesPath, ExcelCellBase.GetFullAddress(_chart.WorkSheet.Name, _xSeries));
                    }
                    else
                    {
                        SetXmlNodeString(_xSeriesPath, _xSeries);
                    }
                    SetXSerieFunction();
                }
            }
       }

        private void GetLitValues(string value, out double[] numberLiterals, out string[] stringLiterals)
        {
            value = value.Substring(1, value.Length - 2); //Remove outer {}
            if (value[0] == '\"' || value[0] == '\'')
            {
                numberLiterals = null;
                stringLiterals = SplitStringValue(value, value[0]);
            }
            else
            {
                stringLiterals = null;
                var split = value.Split(',');
                numberLiterals = new double[split.Length];

                for (int i = 0; i < split.Length; i++)
                {
                    if (double.TryParse(split[i], NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                    {
                        numberLiterals[i] = d;
                    }
                }
            }
        }

        private string[] SplitStringValue(string value, char textQualifier)
        {
            var sb = new StringBuilder();
            bool insideStr = true;
            var list = new List<string>();
            for (int i = 1; i < value.Length; i++)
            {
                if (insideStr)
                {
                    if (value[i] == textQualifier)
                    {
                        insideStr = false;
                    }
                    else
                    {
                        sb.Append(value[i]);
                    }
                }
                else
                {
                    if (value[i] == textQualifier)
                    {
                        insideStr = true;
                        if (sb.Length > 0)
                        {
                            sb.Append(value[i]);
                        }
                    }
                    else if (value[i] == ',')
                    {
                        list.Add(sb.ToString());
                        sb = new StringBuilder();
                    }
                    else
                    {
                        throw (new InvalidOperationException($"String array has an invalid format at position {i}"));
                    }
                }
            }
            if (sb.Length > 0)
            {
                list.Add(sb.ToString());
            }

            return list.ToArray();
        }
        private void SetSerieFunction(string value)
        {
            CreateNode(_seriesPath, true);
            CreateNode(_numCachePath, true);

            SetXmlNodeString(_seriesPath, ToFullAddress(value));

            if (_chart.PivotTableSource != null)
            {
                XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), NameSpaceManager);
                if (cache != null)
                {
                    cache.ParentNode.RemoveChild(cache);
                }
                SetXmlNodeString(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), "General");
            }

            XmlNode lit = TopNode.SelectSingleNode(_seriesNumLitPath, NameSpaceManager);
            if (lit != null)
            {
                lit.ParentNode.RemoveChild(lit);
            }
        }

        private void SetXSerieFunction()
        {
            if (_xSeriesPath.IndexOf("c:numRef", StringComparison.OrdinalIgnoreCase) > 0)
            {
                XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _xSeriesTopPath), NameSpaceManager);
                if (cache != null)
                {
                    cache.ParentNode.RemoveChild(cache);
                }

                XmlNode lit = TopNode.SelectSingleNode(_xSeriesNumLitPath, NameSpaceManager);
                if (lit != null)
                {
                    lit.ParentNode.RemoveChild(lit);
                }
            }
            else
            {
                XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:strRef/c:strCache", _xSeriesTopPath), NameSpaceManager);
                if (cache != null)
                {
                    cache.ParentNode.RemoveChild(cache);
                }

                XmlNode lit = TopNode.SelectSingleNode(_xSeriesStrLitPath, NameSpaceManager);
                if (lit != null)
                {
                    lit.ParentNode.RemoveChild(lit);
                }
            }
        }
        private void SetLits(double[] numLit, string[] strLit, string numLitPath, string strLitPath)
        {
            if(strLit!=null)
            {
                XmlNode lit = CreateNode(strLitPath);
                SetLitArray(lit, strLit);
            }
            else if(numLit!=null)
            {
                XmlNode lit = CreateNode(numLitPath);
                SetLitArray(lit, numLit);
            }
        }

        private void SetLitArray(XmlNode lit, double[] numLit)
        {
            if (numLit.Length == 0) return;
            var ci = CultureInfo.InvariantCulture;
            for (int i = 0; i < numLit.Length; i++)
            {
                var pt = lit.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                pt.SetAttribute("idx", i.ToString(CultureInfo.InvariantCulture));
                lit.AppendChild(pt);
                pt.InnerXml = $"<c:v>{((double)numLit[i]).ToString("R15", ci)}</c:v>";
            }
            AddCount(lit, numLit.Length);
        }

        private void SetLitArray(XmlNode lit, string[] strLit)
        {
            for (int i = 0; i < strLit.Length; i++)
            {
                var pt = lit.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                pt.SetAttribute("idx", i.ToString(CultureInfo.InvariantCulture));
                lit.AppendChild(pt);
                pt.InnerXml = $"<c:v>{strLit[i]}</c:v>";
            }
            AddCount(lit, strLit.Length);
        }
        private static void AddCount(XmlNode lit, int count)
        {
            var ct = lit.OwnerDocument.CreateElement("c", "ptCount", ExcelPackage.schemaChart);
            ct.SetAttribute("val", count.ToString(CultureInfo.InvariantCulture));
            lit.InsertBefore(ct, lit.FirstChild);
        }

        ExcelChartTrendlineCollection _trendLines = null;
       /// <summary>
       /// Access to the trendline collection
       /// </summary>
        public override ExcelChartTrendlineCollection TrendLines
        {
            get
            {
                if (_trendLines == null)
                {
                    _trendLines = new ExcelChartTrendlineCollection(this);
                }
                return _trendLines;
            }
        }
        /// <summary>
        /// Number of items in the serie
        /// </summary>
        public override int NumberOfItems
        {
            get
            {
                if(ExcelCellBase.IsValidAddress(Series))
                {
                    var a = new ExcelAddressBase(Series);
                    return a.Rows;
                }
                else
                {
                    return 30;  //For unhandled sources (tables, pivottables and functions), set the items to 30. This is will generate 30 datapoints for which in most cases are sufficent.
                }
            }
        }

        /// <summary>
        /// Creates a num cach for a chart serie.
        /// Please note that a serie can only have one column to have a cache.        
        /// </summary>
        public void CreateCache()
        {
            if (_isPivot) throw(new NotImplementedException("Cache for pivotcharts has not been implemented yet."));

            if (!string.IsNullOrEmpty(Series))
            {
                if(new ExcelRangeBase(_chart.WorkSheet, Series).Columns > 1)
                {
                    throw (new InvalidOperationException("A serie cannot be multiple columns. Please add one serie per column to create a cache"));
                }
                var node = GetTopNode(Series, _seriesTopPath);
                
                CreateCache(Series, node);
            }

            if (!string.IsNullOrEmpty(XSeries))
            {
                if (new ExcelRangeBase(_chart.WorkSheet, XSeries).Columns > 1)
                {
                    throw (new InvalidOperationException("A serie cannot be multiple columns (XSerie). Please add one serie per column to create a cache"));
                }

                var node = GetTopNode(XSeries, _xSeriesTopPath);

                CreateCache(XSeries, node);
            }
        }
        private void CreateCache(string address, XmlNode node)
        {
            //var ws = _chart.WorkSheet;
            var wb = _chart.WorkSheet.Workbook;
            var addr = new ExcelAddressBase(address);
            if (addr.IsExternal)
            {
                var erIx = wb.ExternalLinks.GetExternalLink(addr._wb);
                if (erIx >= 0 && wb.ExternalLinks[erIx].ExternalLinkType == ExternalReferences.eExternalLinkType.ExternalWorkbook)
                {
                    var er = wb.ExternalLinks[erIx].As.ExternalWorkbook;
                    if (er.Package == null)
                    {
                        CreateCacheFromExternalCache(node, er, addr);
                    }
                    else
                    {
                        CreateCacheFromRange(node, er.Package.Workbook.Worksheets[addr.WorkSheetName]?.Cells[addr.LocalAddress]);
                    }
                }
                else
                {
                    return;
                }
            }
            else
            {
                var ws = string.IsNullOrEmpty(addr.WorkSheetName) ? _chart.WorkSheet : _chart.WorkSheet.Workbook.Worksheets[addr.WorkSheetName];
                if (ws == null) //Worksheet does not exist, exit
                {
                    return;
                }
                CreateCacheFromRange(node, ws.Cells[address]);
            }
            
        }

        private void CreateCacheFromRange(XmlNode node, ExcelRangeBase range)
        {
            if (range == null) return;
            var startRow = range._fromRow;
            var items = 0;
            var cse = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, startRow,range._fromCol, range._toRow, range._toCol);
            while (cse.Next())
            {
                var v = cse.Value._value;
                if (v != null)
                {
                    var d = Utils.ConvertUtil.GetValueDouble(v);
                    var ptNode = node.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                    node.AppendChild(ptNode);
                    ptNode.SetAttribute("idx", (cse.Row - startRow).ToString(CultureInfo.InvariantCulture));
                    ptNode.InnerXml = $"<c:v>{Utils.ConvertUtil.GetValueForXml(d, range.Worksheet.Workbook.Date1904)}</c:v>";
                    items++;
                }
            }

            var countNode = node.SelectSingleNode("c:ptCount", NameSpaceManager) as XmlElement;
            if (countNode != null)
            {
                countNode.SetAttribute("val", items.ToString(CultureInfo.InvariantCulture));
            }
        }
        private void CreateCacheFromExternalCache(XmlNode node, ExternalReferences.ExcelExternalWorkbook er, ExcelAddressBase addr)
        {
            var ews = er.CachedWorksheets[addr.WorkSheetName];
            if (ews == null) return;
            var startRow = addr._fromRow;
            var items = 0;
            var cse = new CellStoreEnumerator<object>(ews.CellValues._values, startRow, addr._fromCol, addr._toRow, addr._toCol);
            while (cse.Next())
            {
                var v = cse.Value;
                if (v != null)
                {
                    var d = Utils.ConvertUtil.GetValueDouble(v);
                    var ptNode = node.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                    node.AppendChild(ptNode);
                    ptNode.SetAttribute("idx", (cse.Row - startRow).ToString(CultureInfo.InvariantCulture));
                    ptNode.InnerXml = $"<c:v>{Utils.ConvertUtil.GetValueForXml(d, er._wb.Date1904)}</c:v>";
                    items++;
                }
            }

            var countNode = node.SelectSingleNode("c:ptCount", NameSpaceManager) as XmlElement;
            if (countNode != null)
            {
                countNode.SetAttribute("val", items.ToString(CultureInfo.InvariantCulture));
            }
        }

        private XmlNode GetTopNode(string address, string seriesTopPath)
        {
            if (ExcelCellBase.IsValidAddress(address))
            {
                var addr = new ExcelAddressBase(address);
                object v;
                var wb = _chart.WorkSheet.Workbook;
                if (addr.IsExternal)
                {
                    var erIx = wb.ExternalLinks.GetExternalLink(addr._wb);
                    if(erIx>=0)
                    {
                        var er = wb.ExternalLinks[erIx].As.ExternalWorkbook;
                        if(er.Package!=null)
                        {
                            var ws = er.Package.Workbook.Worksheets[addr.WorkSheetName];
                            var range = ws.Cells[addr.LocalAddress];
                            v = range.FirstOrDefault()?.Value;
                        }
                        else
                        {
                            var ws = er.CachedWorksheets[addr.WorkSheetName];
                            if(ws==null)
                            {
                                v = null;
                            }
                            else
                            {
                                //Get the first value in the cached range.
                                v = ws.CellValues[addr._fromRow, addr._fromCol];
                            }
                        }
                    }
                    else
                    {
                        v = null;
                    }
                }
                else 
                {
                    ExcelWorksheet ws;
                    if (string.IsNullOrEmpty(addr.WorkSheetName))
                    {
                        ws = _chart.WorkSheet;
                    }
                    else
                    {
                        ws = _chart.WorkSheet.Workbook.Worksheets[addr.WorkSheetName];
                    }
                    if (ws == null)
                    {
                        v = null;
                    }
                    else
                    {
                        var range = ws.Cells[address];
                        v = range.FirstOrDefault()?.Value;
                    }
                }
                

                string cachePath;
                bool isNum;
                if(Utils.ConvertUtil.IsNumericOrDate(v) || v is null)
                {
                    cachePath = string.Format("{0}/c:numRef/c:numCache", seriesTopPath);
                    isNum = true;
                }
                else
                {
                    cachePath=string.Format("{0}/c:strRef/c:strCache", seriesTopPath);
                    isNum = false;
                }
                var node = CreateNode(cachePath);
                if (node.HasChildNodes)
                {
                    if(isNum)
                    {
                        if(node.FirstChild.LocalName== "formatCode")
                        {
                            node.InnerXml = node.FirstChild.OuterXml;
                        }
                        else
                        {
                            node.InnerXml = "";
                        }
                    }
                    else
                    {
                        node.InnerXml = ""; 
                    }
                }
                CreateNode($"{cachePath}/c:ptCount");
                return node;
            }
            else
            {
                throw (new NotImplementedException("Litteral cache has not been implemented yet."));
            }
        }
        internal static XmlElement CreateSerieElement(ExcelChart chart)
        {
            var ser = (XmlElement)chart._chartXmlHelper.CreateNode("c:ser", false, true);

            //If the chart is added from a chart template, then use the chart templates series xml
            if (!string.IsNullOrEmpty(chart._drawings._seriesTemplateXml))
            {
                ser.InnerXml = chart._drawings._seriesTemplateXml;
            }

            int idx = FindIndex(chart._topChart??chart);
            ser.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{2}{5}{0}{3}{4}", AddExplosion(chart.ChartType), idx, AddSpPrAndScatterPoint(chart.ChartType), AddAxisNodes(chart.ChartType), AddSmooth(chart.ChartType), AddMarker(chart.ChartType));
            return ser;
        }

        private static int FindIndex(ExcelChart chart)
        {
            int ret = 0, newID = 0;
            if (chart.PlotArea.ChartTypes.Count > 1)
            {
                foreach (var chartType in chart.PlotArea.ChartTypes)
                {
                    if (newID > 0)
                    {
                        foreach (ExcelChartSerie serie in chartType.Series)
                        {
                            serie.SetID((++newID).ToString());
                        }
                    }
                    else
                    {
                        if (chartType == chart)
                        {
                            ret += chartType.Series.Count + 1;
                            newID = ret;
                        }
                        else
                        {
                            ret += chartType.Series.Count;
                        }
                    }
                }
                return ret - 1;
            }
            else
            {
                return chart.Series.Count;
            }
        }
        #region "Xml init Functions"
        private static string AddMarker(eChartType chartType)
        {
            if (chartType == eChartType.Line ||
                chartType == eChartType.LineStacked ||
                chartType == eChartType.LineStacked100 ||
                chartType == eChartType.XYScatterLines ||
                chartType == eChartType.XYScatterSmooth ||
                chartType == eChartType.XYScatterLinesNoMarkers ||
                chartType == eChartType.XYScatterSmoothNoMarkers)
            {
                return "<c:marker><c:symbol val=\"none\" /></c:marker>";
            }
            else
            {
                return "";
            }
        }
        private static string AddSpPrAndScatterPoint(eChartType chartType)
        {
            if (chartType == eChartType.XYScatter)
            {
                return "<c:spPr><a:noFill/><a:ln w=\"28575\"><a:noFill /></a:ln><a:effectLst/><a:sp3d/></c:spPr>";
            }
            else
            {
                return "";
            }
        }
        private static string AddAxisNodes(eChartType chartType)
        {
            if (chartType == eChartType.XYScatter ||
                 chartType == eChartType.XYScatterLines ||
                 chartType == eChartType.XYScatterLinesNoMarkers ||
                 chartType == eChartType.XYScatterSmooth ||
                 chartType == eChartType.XYScatterSmoothNoMarkers ||
                 chartType == eChartType.Bubble ||
                 chartType == eChartType.Bubble3DEffect)
            {
                return "<c:xVal /><c:yVal />";
            }
            else
            {
                return "<c:val />";
            }
        }

        private static string AddExplosion(eChartType chartType)
        {
            if (chartType == eChartType.PieExploded3D ||
               chartType == eChartType.PieExploded ||
                chartType == eChartType.DoughnutExploded)
            {
                return "<c:explosion val=\"25\" />"; //Default 25;
            }
            else
            {
                return "";
            }
        }
        private static string AddSmooth(eChartType chartType)
        {
            if (chartType == eChartType.XYScatterSmooth ||
               chartType == eChartType.XYScatterSmoothNoMarkers)
            {
                return "<c:smooth val=\"1\" />"; //Default 25;
            }
            else
            {
                return "";
            }
        }
        #endregion
    }
}
