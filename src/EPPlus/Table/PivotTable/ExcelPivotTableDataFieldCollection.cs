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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Collection class for data fields in a Pivottable 
    /// </summary>
    public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
    {
        private readonly ExcelPivotTable _table;

        internal ExcelPivotTableDataFieldCollection(ExcelPivotTable table) :
            base()
        {
            _table = table;
        }
        /// <summary>
        /// Add a new datafield
        /// </summary>
        /// <param name="field">The field</param>
        /// <returns>The new datafield</returns>
        public ExcelPivotTableDataField Add(ExcelPivotTableField field)
        {
            if(field==null)
            {
                throw new ArgumentNullException("field", "Parameter field cannot be null");
            }
            var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            if (dataFieldsNode == null)
            {
                _table.CreateNode("d:dataFields");
                dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            }

            XmlElement node = _table.PivotTableXml.CreateElement("dataField", ExcelPackage.schemaMain);
            node.SetAttribute("fld", field.Index.ToString());
            dataFieldsNode.AppendChild(node);

            //XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
            field.SetXmlNodeBool("@dataField", true,false);

            var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, node, field);
            ValidateDupName(dataField);

            _list.Add(dataField);
            return dataField;
        }
        private void ValidateDupName(ExcelPivotTableDataField dataField)
        {
            if(ExistsDfName(dataField.Field.Name, null))
            {
                var index = 2;
                string name;
                do
                {
                    name = dataField.Field.Name + "_" + index++.ToString();
                }
                while (ExistsDfName(name,null));
                dataField.Name = name;
            }
        }

        internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
        {
            foreach (var df in _list)
            {
                if (((!string.IsNullOrEmpty(df.Name) && df.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                     (string.IsNullOrEmpty(df.Name) && df.Field.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))) && datafield != df)
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Remove a datafield
        /// </summary>
        /// <param name="dataField">The data field to remove</param>
        public void Remove(ExcelPivotTableDataField dataField)
        {
            XmlElement node = dataField.Field.TopNode.SelectSingleNode(string.Format("../../d:dataFields/d:dataField[@fld={0}]", dataField.Index), dataField.NameSpaceManager) as XmlElement;
            if (node != null)
            {
                node.ParentNode.RemoveChild(node);
            }
            _list.Remove(dataField);
        }
    }
}