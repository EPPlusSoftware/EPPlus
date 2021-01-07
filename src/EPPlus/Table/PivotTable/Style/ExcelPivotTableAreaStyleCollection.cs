using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableAreaStyleCollection : EPPlusReadOnlyList<ExcelPivotTableAreaStyle>
    {
        ExcelStyles _styles;
        XmlHelper _xmlHelper;
        ExcelPivotTable _pt;
        internal ExcelPivotTableAreaStyleCollection(ExcelPivotTable pt)
        {
            _pt = pt;
            _styles = pt.WorkSheet.Workbook.Styles;
            foreach (XmlNode node in pt.GetNodes("d:formats/d:format/d:pivotArea"))
            {
                var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, node, _styles);
            }
        }
        internal ExcelPivotTableAreaStyle Add()
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _styles);
            _list.Add(s);
            return s;
        }

        internal ExcelPivotTableAreaStyle Add(ePivotAreaType type)
        {
            var formatNode = GetTopNode();
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _styles)
            {
                PivotAreaType = type
            };
            _list.Add(s);
            return s;
        }

        internal ExcelPivotTableAreaStyle Add(ePivotAreaType type, ePivotTableAxis axis)
        {
            var formatNode = GetTopNode();
            
            var s = new ExcelPivotTableAreaStyle(_styles.NameSpaceManager, formatNode.FirstChild, _styles)
            {
                PivotAreaType = type,
                Axis = axis
            };
            _list.Add(s);
            return s;
        }
        private XmlNode GetTopNode()
        {
            if (_xmlHelper == null)
            {
                var node = _pt.CreateNode("d:formats");
                _xmlHelper = XmlHelperFactory.Create(_pt.NameSpaceManager, node);
            }
            
            var retNode = _xmlHelper.CreateNode("d:format",false,true);
            retNode.InnerXml = $"<pivotArea xmlns=\"{ExcelPackage.schemaMain}\"/>";
            return retNode;
        }
    }
}
