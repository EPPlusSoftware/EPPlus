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
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartSerie : XmlHelper, IDrawingStyleBase
   {
        //internal ExcelChartSeries _chartSeries;
        internal ExcelChart _chart;
        private readonly bool _isPivot;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
       /// <param name="isPivot">Is pivotchart</param>  
       internal ExcelChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(ns,node)
       {
            //_chartSeries = chartSeries;
            _chart = chart;
           _isPivot = isPivot;
           SchemaNodeOrder = new string[] { "idx", "order", "tx", "spPr", "marker", "invertIfNegative", "pictureOptions", "explosion", "dPt", "dLbls", "trendline","errBars", "cat", "val", "xVal", "yVal", "smooth","shape", "bubbleSize", "bubble3D", "numRef", "numLit", "strRef", "strLit", "formatCode", "ptCount", "pt" };

           if (_chart.ChartType == eChartType.XYScatter ||
               _chart.ChartType == eChartType.XYScatterLines ||
               _chart.ChartType == eChartType.XYScatterLinesNoMarkers ||
               _chart.ChartType == eChartType.XYScatterSmooth ||
               _chart.ChartType == eChartType.XYScatterSmoothNoMarkers ||
               _chart.ChartType == eChartType.Bubble ||
               _chart.ChartType == eChartType.Bubble3DEffect)
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

            if(ExistNode(sp))
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
        internal void SetID(string id)
       {
           SetXmlNodeString("c:idx/@val",id);
           SetXmlNodeString("c:order/@val", id);
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public string Header 
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
       public ExcelAddressBase HeaderAddress
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
                SetXmlNodeString(headerAddressPath, ExcelCellBase.GetFullAddress(value.WorkSheet, value.Address));
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
        public virtual string Series
       {
           get
           {
               return GetXmlNodeString(_seriesPath);
           }
           set
           {
                value = value.Trim();
                if (value.StartsWith("=")) value = value.Substring(1);
                if (value.StartsWith("{") && value.EndsWith("}"))
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
        public virtual string XSeries
       {
           get
           {
               return GetXmlNodeString(_xSeriesPath);
           }
           set
           {
                _xSeries = value.Trim();
                if (_xSeries.StartsWith("=")) _xSeries = _xSeries.Substring(1);
                if (value.StartsWith("{") && value.EndsWith("}"))
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

        /// <summary>
        /// Literals for the Y serie, if the literal values are numeric
        /// </summary>
        public double[] NumberLiteralsY { get; private set; } = null;
        /// <summary>
        /// Literals for the X serie, if the literal values are numeric
        /// </summary>
        public double[] NumberLiteralsX { get; private set; } = null;
        /// <summary>
        /// Literals for the X serie, if the literal values are strings
        /// </summary>
        public string[] StringLiteralsX { get; private set; } = null;

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
            
            if(ExcelCellBase.IsValidAddress(value))
            {
                SetXmlNodeString(_seriesPath, ExcelCellBase.GetFullAddress(_chart.WorkSheet.Name, value));
            }
            else
            {
                SetXmlNodeString(_seriesPath, value);
            }

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
            if (_xSeriesPath.IndexOf("c:numRef") > 0)
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
        public ExcelChartTrendlineCollection TrendLines
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
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }
        /// <summary>
        /// Number of items in the serie
        /// </summary>
        public int NumberOfItems
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
                    throw (new InvalidOperationException("A serie can not be multiple columns. Please add one serie per column to create a cache"));
                }
                var node = GetTopNode(Series, _seriesTopPath);
                
                CreateCache(Series, node);
            }

            if (!string.IsNullOrEmpty(XSeries))
            {
                if (new ExcelRangeBase(_chart.WorkSheet, XSeries).Columns > 1)
                {
                    throw (new InvalidOperationException("A serie can not be multiple columns (XSerie). Please add one serie per column to create a cache"));
                }

                var node = GetTopNode(XSeries, _xSeriesTopPath);

                CreateCache(XSeries, node);
            }
        }
        private void CreateCache(string address, XmlNode node)
        {
            var ws = _chart.WorkSheet;
            var range = ws.Cells[address];
            var startRow = range._fromRow;
            var items = 0;
            var cse = new CellStoreEnumerator<ExcelValue>(ws._values);
            while(cse.Next())
            {
                var v = cse.Value._value;
                if (v != null)
                {
                    var d = Utils.ConvertUtil.GetValueDouble(v);
                    var ptNode = node.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                    node.AppendChild(ptNode);
                    ptNode.SetAttribute("idx", (cse.Row - startRow).ToString(CultureInfo.InvariantCulture));
                    ptNode.InnerXml = $"<c:v>{Utils.ConvertUtil.GetValueForXml(d, ws.Workbook.Date1904)}</c:v>";
                    items++;
                }                
            }

            var countNode = node.SelectSingleNode("c:ptCount", NameSpaceManager) as XmlElement;
            if(countNode != null)
            {
                countNode.SetAttribute("val", items.ToString(CultureInfo.InvariantCulture));
            }
        }

        private XmlNode GetTopNode(string address, string seriesTopPath)
        {
            if (ExcelCellBase.IsValidAddress(address))
            {
                var ws = _chart.WorkSheet;
                var range = ws.Cells[address];
                var v = range.FirstOrDefault()?.Value;

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

    }
}
