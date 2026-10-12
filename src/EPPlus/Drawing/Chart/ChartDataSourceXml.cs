//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;
//using System.Text;
//using System.Xml;

//namespace OfficeOpenXml.Drawing.Chart
//{
//    internal class ChartDataSourceXml : XmlHelper
//    {
//    //    string _seriesPath;
//    //    string _numCachePath;

//    //    internal ChartDataSourceXml(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
//    //    {

//    //    }

//    //    private void SetSerieFunction(string value)
//    //    {
//    //        CreateNode(_seriesPath, true);
//    //        CreateNode(_numCachePath, true);

//    //        SetXmlNodeString(_seriesPath, ToFullAddress(value));

//    //        if (_chart.PivotTableSource != null)
//    //        {
//    //            XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), NameSpaceManager);
//    //            if (cache != null)
//    //            {
//    //                cache.ParentNode.RemoveChild(cache);
//    //            }
//    //            SetXmlNodeString(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), "General");
//    //        }

//    //        XmlNode lit = TopNode.SelectSingleNode(_seriesNumLitPath, NameSpaceManager);
//    //        if (lit != null)
//    //        {
//    //            lit.ParentNode.RemoveChild(lit);
//    //        }
//    //    }

//    //    private void SetXSerieFunction()
//    //    {
//    //        if (_xSeriesPath.IndexOf("c:numRef", StringComparison.OrdinalIgnoreCase) > 0)
//    //        {
//    //            XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache", _xSeriesTopPath), NameSpaceManager);
//    //            if (cache != null)
//    //            {
//    //                cache.ParentNode.RemoveChild(cache);
//    //            }

//    //            XmlNode lit = TopNode.SelectSingleNode(_xSeriesNumLitPath, NameSpaceManager);
//    //            if (lit != null)
//    //            {
//    //                lit.ParentNode.RemoveChild(lit);
//    //            }
//    //        }
//    //        else
//    //        {
//    //            XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:strRef/c:strCache", _xSeriesTopPath), NameSpaceManager);
//    //            if (cache != null)
//    //            {
//    //                cache.ParentNode.RemoveChild(cache);
//    //            }

//    //            XmlNode lit = TopNode.SelectSingleNode(_xSeriesStrLitPath, NameSpaceManager);
//    //            if (lit != null)
//    //            {
//    //                lit.ParentNode.RemoveChild(lit);
//    //            }
//    //        }
//    //    }

//    //    private void SetLits(double[] numLit, string[] strLit, string numLitPath, string strLitPath)
//    //    {
//    //        if (strLit != null)
//    //        {
//    //            XmlNode lit = CreateNode(strLitPath);
//    //            SetLitArray(lit, strLit);
//    //        }
//    //        else if (numLit != null)
//    //        {
//    //            XmlNode lit = CreateNode(numLitPath);
//    //            SetLitArray(lit, numLit);
//    //        }
//    //    }

//    //    private void SetLitArray(XmlNode lit, double[] numLit)
//    //    {
//    //        if (numLit.Length == 0) return;
//    //        var ci = CultureInfo.InvariantCulture;

//    //        //Remove previous child nodes
//    //        var previousPt = lit.SelectNodes("c:pt", NameSpaceManager);
//    //        if (previousPt != null)
//    //        {
//    //            for (int i = 0; i < previousPt.Count; i++)
//    //            {
//    //                lit.RemoveChild(previousPt[i]);
//    //            }
//    //        }

//    //        for (int i = 0; i < numLit.Length; i++)
//    //        {
//    //            var pt = lit.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
//    //            pt.SetAttribute("idx", i.ToString(CultureInfo.InvariantCulture));
//    //            lit.AppendChild(pt);
//    //            pt.InnerXml = $"<c:v>{((double)numLit[i]).ToString("R15", ci)}</c:v>";
//    //        }
//    //        AddCount(lit, numLit.Length);
//    //    }

//    //    private void SetLitArray(XmlNode lit, string[] strLit)
//    //    {
//    //        //Remove previous child nodes
//    //        var previousPt = lit.SelectNodes("c:pt", NameSpaceManager);
//    //        if (previousPt != null)
//    //        {
//    //            for (int i = 0; i < previousPt.Count; i++)
//    //            {
//    //                lit.RemoveChild(previousPt[i]);
//    //            }
//    //        }

//    //        for (int i = 0; i < strLit.Length; i++)
//    //        {
//    //            var pt = lit.OwnerDocument.CreateElement("c", "pt", ExcelPackage.schemaChart);
//    //            pt.SetAttribute("idx", i.ToString(CultureInfo.InvariantCulture));
//    //            lit.AppendChild(pt);
//    //            pt.InnerXml = $"<c:v>{strLit[i]}</c:v>";
//    //        }
//    //        AddCount(lit, strLit.Length);
//    //    }
//    //    private void AddCount(XmlNode lit, int count)
//    //    {
//    //        var ct = (XmlElement)lit.SelectSingleNode("c:ptCount", NameSpaceManager);
//    //        if (ct == null)
//    //        {
//    //            ct = lit.OwnerDocument.CreateElement("c", "ptCount", ExcelPackage.schemaChart);
//    //            lit.InsertBefore(ct, lit.FirstChild);
//    //        }
//    //        ct.SetAttribute("val", count.ToString(CultureInfo.InvariantCulture));
//    //    }
//    }
//}
