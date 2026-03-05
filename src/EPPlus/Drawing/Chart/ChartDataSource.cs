using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Base-class for things like Series, XSeries, DataLabelsRange
    /// This class and children should loosely represent "CT_AxDataSource" in dml-chart.xsd
    /// </summary>
    internal class ChartDataSource : XmlHelper
    {
        private readonly bool _isPivot;
        ExcelWorksheet _ws;

        private bool _isExt = false;

        string _strRef = null;

        string sourceTopPath;
        string _StrRefParentPath = "{0}/{1}";
        string _StrRefPath = "{0}/{1}/c:f";
        string _xSeriesStrLitPath, _xSeriesNumLitPath;

        internal string StrRefFormulaValue
        {
            get { return _strRef; }
        }

        internal bool RefIsValidAddress { private set; get; } = false;

        private double[] _numberLiterals;
        internal double[] NumberLiterals
        {
            get
            {
                if (string.IsNullOrEmpty(_xSeriesNumLitPath) == false && GetNode(_xSeriesNumLitPath) != null)
                {
                    ReadNumLiterals(_xSeriesNumLitPath, out _numberLiterals);
                    return _numberLiterals;
                }
                return _numberLiterals;
            }
            set
            {
                _numberLiterals = value;
            }
        }

        private string[] _stringLiterals;
        internal string[] StringLiterals
        {
            get
            {
                if (string.IsNullOrEmpty(_xSeriesStrLitPath) == false && GetNode(_xSeriesStrLitPath) != null)
                {
                    ReadStringLiterals(_xSeriesStrLitPath, out _stringLiterals);
                    return _stringLiterals;
                }
                return _stringLiterals;
            }
            set
            {
                _stringLiterals = value;
            }
        }

        /// <summary>
        /// Base-class for things like Series, XSeries, DataLabelsRange
        /// This class and children should loosely represent "CT_AxDataSource" in dml-chart.xsd
        /// </summary>
        /// <param name="isPivot">If this source is part of a pivotTable</param>
        ///
        internal ChartDataSource(bool isPivot, XmlNamespaceManager ns, XmlNode node, ExcelWorksheet ws, string seriesTopPath) : base(ns, node)
        {
            _isPivot = isPivot;
            _ws = ws;
            sourceTopPath = seriesTopPath;

            var ep = string.Format(_StrRefParentPath, sourceTopPath, "c15:datalabelsRange");

            if (ExistsNode(ep))
            {
                _isExt = true;
                _StrRefPath = ep + "/c15:f";
                _xSeriesStrLitPath = ep + "/c15:dlblRangeCache";
            }
        }

        internal void SetStrRef(string strRef)
        {
            _strRef = strRef.Trim();
            if (_strRef.StartsWith("=", StringComparison.OrdinalIgnoreCase)) _strRef = _strRef.Substring(1);

            if (strRef.StartsWith("{", StringComparison.OrdinalIgnoreCase) && strRef.EndsWith("}", StringComparison.OrdinalIgnoreCase))
            {
                if(_isExt)
                {
                    CreateNode(_StrRefPath, true);
                    SetXmlNodeString(_StrRefPath, _strRef);
                }

                GetLitValues(_strRef, out double[] numLit, out string[] strLit);
                NumberLiterals = numLit;
                StringLiterals = strLit;
                SetLits(numLit, strLit, _xSeriesNumLitPath, _xSeriesStrLitPath);
            }
            else
            {
                NumberLiterals = null;
                StringLiterals = null;
                CreateNode(_StrRefPath, true);

                if (ExcelCellBase.IsValidAddress(strRef))
                {
                    RefIsValidAddress = true;
                    SetXmlNodeString(_StrRefPath, ExcelCellBase.GetFullAddress(_ws.Name, _strRef));
                }
                else
                {
                    SetXmlNodeString(_StrRefPath, _strRef);
                }
                SetSourceFunction();
            }
        }

        private void SetLits(double[] numLit, string[] strLit, string numLitPath, string strLitPath)
        {
            if (strLit != null)
            {
                XmlNode lit = CreateNode(strLitPath);
                SetLitArray(lit, strLit);
            }
            else if (numLit != null)
            {
                XmlNode lit = CreateNode(numLitPath);
                SetLitArray(lit, numLit);
            }
        }


        private void ReadNumLiterals(string path, out double[] numberLiterals)
        {
            var childNodes = GetNode(path).ChildNodes;
            numberLiterals = new double[childNodes.Count];
            List<double> numLits = new();

            foreach (XmlNode node in childNodes)
            {
                if (node.NodeType == XmlNodeType.Element && node.LocalName == "pt")
                {
                    if (double.TryParse(node.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out double numLit) == false)
                    {
                        throw new InvalidDataException($"numberLiteral in xml node:'{node.Name}' in ws:'{_ws.Name}' with value:'{node.InnerText}' could not be parsed as double. Chart cannot be read.");
                    }
                    numLits.Add(numLit);
                }
            }
            numberLiterals = numLits.ToArray();
        }

        private void SetLitArray(XmlNode lit, double[] numLit)
        {
            if (numLit.Length == 0) return;
            var ci = CultureInfo.InvariantCulture;

            //Remove previous child nodes
            var previousPt = lit.SelectNodes("c:pt", NameSpaceManager);
            if (previousPt != null)
            {
                for (int i = 0; i < previousPt.Count; i++)
                {
                    lit.RemoveChild(previousPt[i]);
                }
            }

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
            //Remove previous child nodes
            var previousPt = lit.SelectNodes("c:pt", NameSpaceManager);
            if (previousPt != null)
            {
                for (int i = 0; i < previousPt.Count; i++)
                {
                    lit.RemoveChild(previousPt[i]);
                }
            }

            for (int i = 0; i < strLit.Length; i++)
            {
                var pt = lit.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                pt.SetAttribute("idx", i.ToString(CultureInfo.InvariantCulture));
                lit.AppendChild(pt);
                pt.InnerXml = $"<c:v>{strLit[i]}</c:v>";
            }
            AddCount(lit, strLit.Length);
        }
        private void AddCount(XmlNode lit, int count)
        {
            var ct = (XmlElement)lit.SelectSingleNode("c:ptCount", NameSpaceManager);
            if (ct == null)
            {
                ct = lit.OwnerDocument.CreateElement("c", "ptCount", ExcelPackage.schemaChart);
                lit.InsertBefore(ct, lit.FirstChild);
            }
            ct.SetAttribute("val", count.ToString(CultureInfo.InvariantCulture));
        }

        private void ReadStringLiterals(string path, out string[] stringLiterals)
        {
            var parentNode = GetNode(path);
            List<string> strLits = new();

            if (parentNode != null)
            {
                var childNodes = parentNode.ChildNodes;

                foreach (XmlNode node in childNodes)
                {
                    if (node.NodeType == XmlNodeType.Element && node.LocalName == "pt")
                    {
                        strLits.Add(node.InnerText);
                    }
                }
            }
            stringLiterals = strLits.ToArray();
        }

        private void SetSourceFunction()
        {
            if (_StrRefPath.IndexOf("c:numRef", StringComparison.OrdinalIgnoreCase) > 0)
            {
                XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", sourceTopPath), NameSpaceManager);
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
                XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:strRef/c:strCache", sourceTopPath), NameSpaceManager);
                if (cache != null)
                {
                    cache.ParentNode.RemoveChild(cache);
                }

                XmlNode lit = TopNode.SelectSingleNode(_xSeriesStrLitPath, NameSpaceManager);
                if (lit != null)
                {
                    lit.ParentNode.RemoveChild(lit);
                }

                var extCacheStr = string.Format("{0}/c15:datalabelsRange/c15:dlblRangeCache", sourceTopPath);
                XmlNode extCache = TopNode.SelectSingleNode(extCacheStr, NameSpaceManager);
                if (extCache != null)
                {
                    extCache.ParentNode.RemoveChild(extCache);
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

        /// <summary>
        /// Creates a num cach for a chart serie.
        /// Please note that a serie can only have one column to have a cache.        
        /// </summary>
        /// should be public later
        internal void CreateCache(XmlNode seriesTopNode)
        {
            if (_isPivot) throw (new NotImplementedException("Cache for pivotcharts has not been implemented yet."));

            if (!string.IsNullOrEmpty(StrRefFormulaValue))
            {
                var addr = new ExcelRangeBase(_ws, StrRefFormulaValue);
                bool moreThanOneRow = addr.Rows > 1;
                bool moreThanOneColumn = addr.Columns > 1;

                if (moreThanOneColumn)
                {
                    throw (new InvalidOperationException("A serie cannot be multiple columns. Please add one serie per column to create a cache"));
                }

                CreateCache(StrRefFormulaValue, seriesTopNode);
            }
        }
        internal void CreateCache(string address, XmlNode node)
        {
            //var ws = _chart.WorkSheet;
            var wb = _ws.Workbook;
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
                var ws = string.IsNullOrEmpty(addr.WorkSheetName) ? _ws : _ws.Workbook.Worksheets[addr.WorkSheetName];
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
            lastCachedValues.Clear();
            var startRow = range._fromRow;
            var items = 0;
            var cse = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, startRow, range._fromCol, range._toRow, range._toCol);
            while (cse.Next())
            {
                var v = cse.Value._value;
                if (v != null)
                {
                    string xmlValue = "";
                    if (v.IsNumeric())
                    {
                        var d = Utils.TypeConversion.ConvertUtil.GetValueDouble(v);
                        xmlValue = Utils.TypeConversion.ConvertUtil.GetValueForXml(d, range.Worksheet.Workbook.Date1904);
                    }
                    else
                    {
                        xmlValue = string.Format(CultureInfo.InvariantCulture, v.ToString());
                    }

                    var ptNode = node.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                    node.AppendChild(ptNode);
                    ptNode.SetAttribute("idx", (cse.Row - startRow).ToString(CultureInfo.InvariantCulture));
                    lastCachedValues.Add(xmlValue);
                    ptNode.InnerXml = $"<c:v>{xmlValue}</c:v>";
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
                    var d = Utils.TypeConversion.ConvertUtil.GetValueDouble(v);
                    var ptNode = node.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
                    node.AppendChild(ptNode);
                    ptNode.SetAttribute("idx", (cse.Row - startRow).ToString(CultureInfo.InvariantCulture));
                    var xmlValue = Utils.TypeConversion.ConvertUtil.GetValueForXml(d, er._wb.Date1904);
                    lastCachedValues.Add(xmlValue);
                    ptNode.InnerXml = $"<c:v>{xmlValue}</c:v>";
                    items++;
                }
            }

            var countNode = node.SelectSingleNode("c:ptCount", NameSpaceManager) as XmlElement;
            if (countNode != null)
            {
                countNode.SetAttribute("val", items.ToString(CultureInfo.InvariantCulture));
            }
        }

        private List<string> lastCachedValues = new List<string>();
        internal List<string> GetCachedValues()
        {
            return lastCachedValues;
        }
    }
}
