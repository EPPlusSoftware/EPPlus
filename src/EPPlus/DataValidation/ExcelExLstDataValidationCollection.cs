using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    public class ExcelExLstDataValidationCollection : XmlHelper, IEnumerable<IExcelDataValidation>
    {
        private List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();
        private ExcelWorksheet _worksheet = null;
        private const string ExternalDataValidationPath = "//d:extLst/d:ext/x14:dataValidations";
        private readonly string ExternalDataValidationItemsPath = string.Format("{0}//x14:dataValidation", ExternalDataValidationPath);

        internal ExcelExLstDataValidationCollection(ExcelWorksheet worksheet)
            : base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            _worksheet = worksheet;
            //SchemaNodeOrder = worksheet.SchemaNodeOrder;
            SchemaNodeOrder = new string[]
            {
                "xmlns:x14",
                "uri"
            };

            // check validations in the extLst
            var extLstValidationNodes = worksheet.WorksheetXml.SelectNodes(ExternalDataValidationItemsPath, worksheet.NameSpaceManager);
            if (extLstValidationNodes != null && extLstValidationNodes.Count > 0)
            {
                foreach (XmlNode node in extLstValidationNodes)
                {
                    var address = base.GetXmlNodeString(node, "xm:sqref");
                    var uid = node.Attributes["xr:uid"].Value;
                    var typeSchema = node.Attributes["type"] != null ? node.Attributes["type"].Value : "";
                    var type = ExcelDataValidationType.GetBySchemaName(typeSchema);
                    var val = ExcelDataValidationFactory.Create(type, worksheet, address, node, InternalValidationType.ExtLst, uid);
                    val.Uid = uid;
                    _validations.Add(val);
                }
            }

            if (_validations.Count > 0)
            {
                OnValidationCountChanged();
            }
        }

        private void EnsureRootElementExists()
        {
            var node = TopNode.SelectSingleNode(ExternalDataValidationPath, _worksheet.NameSpaceManager) as XmlElement;
            if (node == null)
            {
                node = (XmlElement)CreateNode(ExternalDataValidationPath.TrimStart('/'), false, true);
                node.SetAttribute("xmlns:xm", ExcelPackage.schemaMainXm);
                ((XmlElement)node.ParentNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);
                ((XmlElement)node.ParentNode).SetAttribute("uri", "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}");
            }
            TopNode = node;
        }

        private void OnValidationCountChanged()
        {

            //if (TopNode != null)
            //{
            //    SetXmlNodeString("@count", _validations.Count.ToString());
            //}
            var dvNode = GetRootNode();
            if (_validations.Count == 0)
            {
                if (dvNode != null)
                {
                    _worksheet.WorksheetXml.DocumentElement.RemoveChild(dvNode);
                }
                _worksheet.ClearValidations();
            }
            else
            {
                var attr = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(ExternalDataValidationPath + "[@count]", _worksheet.NameSpaceManager);
                if (attr == null)
                {
                    dvNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
                }
                dvNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
            }
        }

        internal XmlNode GetRootNode()
        {
            EnsureRootElementExists();
            return TopNode;
        }

        internal void Clear()
        {
            DeleteAllNode(ExternalDataValidationPath.TrimStart('/'));
            _validations.Clear();
        }

        public IExcelDataValidationWithFormula<T> AddValidation<T>(IExcelDataValidationWithFormula<T> item)
            where T : IExcelDataValidationFormula
        {
            EnsureRootElementExists();
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        public IExcelDataValidationWithFormula2<T> AddValidation<T>(IExcelDataValidationWithFormula2<T> item)
            where T : IExcelDataValidationFormula
        {
            EnsureRootElementExists();
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        public bool IsEmpty 
        { 
            get
            {
                return _validations == null || _validations.Count == 0;
            } 
        }

        public IEnumerator<IExcelDataValidation> GetEnumerator()
        {
            return _validations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _validations.GetEnumerator();
        }
    }
}
