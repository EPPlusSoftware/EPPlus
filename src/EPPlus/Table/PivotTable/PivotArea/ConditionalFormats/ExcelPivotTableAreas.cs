using System;
using System.Xml;
using OfficeOpenXml.Core;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;
namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableAreas : EPPlusReadOnlyList<ExcelPivotTableAreaStyle>
    {
        readonly ExcelPivotTable _pt;
        readonly XmlNamespaceManager _nsm;
        readonly XmlNode _topNode;
        readonly XmlHelper _helper;
        internal ExcelPivotTableAreas(ExcelPivotTable pt, XmlNode topNode)
        {
            _pt = pt;
            _nsm = pt.NameSpaceManager;
            _helper = XmlHelperFactory.Create(_nsm, topNode);
            foreach(XmlNode node in topNode.ChildNodes)
            {
                if(node.NodeType==XmlNodeType.Element && node.LocalName == "pivotArea" )
                {
                    var area = new ExcelPivotTableAreaStyle(_nsm, node, _pt);
                    _list.Add(area);
                }
            }
        }
        /// <summary>
        /// Adds a new area for the one or more data fields
        /// </summary>
        /// <param name="fields">The data field(s) where the conditional formatting should be applied. If no fields are supplied all the pivot tables data fields will be added to the area</param>
        /// <returns>The pivot area for the conditional formatting</returns>
        public ExcelPivotTableAreaStyle Add(params ExcelPivotTableDataField[] fields)
        {
            var node = _helper.CreateNode("d:pivotArea", false, true);
            var area = new ExcelPivotTableAreaStyle(_nsm, node, _pt);
            if (fields == null && fields.Length > 0)
            {
                foreach (var field in fields)
                {
                    area.Conditions.DataFields.Add(field);
                }
            }
            else
            {
                if (_pt.DataFields.Count == 0)
                {
                    throw (new InvalidOperationException("Can't add a conditional format to a pivot table with no data fields."));
                }
                foreach (var field in _pt.DataFields)
                {
                    area.Conditions.DataFields.Add(field);
                }
            }

            _list.Add(area);
            return area;
        }
    }
}
