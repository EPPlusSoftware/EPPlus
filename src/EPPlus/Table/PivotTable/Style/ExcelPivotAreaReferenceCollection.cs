/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Core;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotAreaReferenceCollection : EPPlusReadOnlyList<ExcelPivotAreaReference>
    {
        XmlHelper _xmlHelper;
        ExcelPivotTable _pt;
        public ExcelPivotAreaReferenceCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt)
        {
            _xmlHelper = XmlHelperFactory.Create(nsm, topNode);
            _pt = pt;
            foreach (XmlNode n in topNode.ChildNodes)
            {
                if (n.LocalName == "reference")
                {
                    _list.Add(new ExcelPivotAreaReference(nsm, n, pt));
                }
            }
        }
        public ExcelPivotAreaReference Add(ExcelPivotTableField field)
        {
            return Add(field._pivotTable, field.Index);
        }
        public ExcelPivotAreaReference Add(ExcelPivotTable pivotTable, int fieldIndex)
        {
            var n = _xmlHelper.CreateNode("d:references");
            var rn=_xmlHelper.CreateNode(n, "d:reference", true);
            var r =new ExcelPivotAreaReference(_xmlHelper.NameSpaceManager, rn, pivotTable, fieldIndex);            
            _list.Add(r);
            return r;
        }

    }
}