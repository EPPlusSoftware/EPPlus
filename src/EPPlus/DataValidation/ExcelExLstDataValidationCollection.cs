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

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validations for the ext lst (x14) section
    /// </summary>
    //public class ExcelExLstDataValidationCollection : XmlHelper, IEnumerable<IExcelDataValidation>
    //{
    //    private List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();
    //    private ExcelWorksheet _worksheet = null;
    //    private readonly DataValidationFormulaListener _formulaListener = null;
    //    private const string ExternalDataValidationPath = "//d:extLst/d:ext/x14:dataValidations";
    //    private readonly string ExternalDataValidationItemsPath = string.Format("{0}//x14:dataValidation", ExternalDataValidationPath);

    //    internal ExcelExLstDataValidationCollection(ExcelWorksheet worksheet, DataValidationFormulaListener formulaListener)
    //        : base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
    //    {
    //        Require.Argument(worksheet).IsNotNull("worksheet");
    //        _worksheet = worksheet;
    //        _formulaListener = formulaListener;
    //        //SchemaNodeOrder = worksheet.SchemaNodeOrder;
    //        SchemaNodeOrder = new string[]
    //        {
    //            "xmlns:x14",
    //            "uri"
    //        };

    //        // check validations in the extLst
    //        var extLstValidationNodes = worksheet.WorksheetXml.SelectNodes(ExternalDataValidationItemsPath, worksheet.NameSpaceManager);
    //        if (extLstValidationNodes != null && extLstValidationNodes.Count > 0)
    //        {
    //            foreach (XmlNode node in extLstValidationNodes)
    //            {
    //                var address = base.GetXmlNodeString(node, "xm:sqref");
    //                var uid = node.Attributes["xr:uid"] != null && !string.IsNullOrEmpty(node.Attributes["xr:uid"].Value) ? node.Attributes["xr:uid"].Value : ExcelDataValidation.NewId();
    //                var typeSchema = node.Attributes["type"] != null ? node.Attributes["type"].Value : "";
    //                var type = ExcelDataValidationType.GetBySchemaName(typeSchema);
    //                var val = ExcelDataValidationFactory.Create(type, worksheet, address, node, InternalValidationType.ExtLst, uid);
    //                val.RegisterFormulaListener(_formulaListener);
    //                val.Uid = uid;
    //                _validations.Add(val);
    //            }
    //        }

    //        if (_validations.Count > 0)
    //        {
    //            OnValidationCountChanged();
    //        }
    //    }

    //    private void EnsureRootElementExists()
    //    {
    //        var node = TopNode.SelectSingleNode(ExternalDataValidationPath, _worksheet.NameSpaceManager) as XmlElement;
    //        if (node == null)
    //        {
    //            node = (XmlElement)CreateNode(ExternalDataValidationPath.TrimStart('/'), false, true);
    //            node.SetAttribute("xmlns:xm", ExcelPackage.schemaMainXm);
    //            ((XmlElement)node.ParentNode).SetAttribute("xmlns:x14", ExcelPackage.schemaMainX14);
    //            ((XmlElement)node.ParentNode).SetAttribute("uri", "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}");
    //        }
    //        TopNode = node;
    //    }

    //    private void OnValidationCountChanged()
    //    {

    //        //if (TopNode != null)
    //        //{
    //        //    SetXmlNodeString("@count", _validations.Count.ToString());
    //        //}
    //        var dvNode = GetRootNode();
    //        if (_validations.Count == 0)
    //        {
    //            if (dvNode != null)
    //            {
    //                var extNode = dvNode.ParentNode;
    //                extNode.RemoveChild(dvNode);
    //                var x14dvNode = extNode.ParentNode;
    //                x14dvNode.RemoveChild(extNode);
    //                _worksheet.WorksheetXml.DocumentElement.RemoveChild(x14dvNode);
    //            }
    //            _worksheet.ClearValidations();
    //        }
    //        else
    //        {
    //            var attr = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(ExternalDataValidationPath + "[@count]", _worksheet.NameSpaceManager);
    //            if (attr == null)
    //            {
    //                dvNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
    //            }
    //            dvNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
    //        }
    //    }

    //    internal XmlNode GetRootNode()
    //    {
    //        EnsureRootElementExists();
    //        return TopNode;
    //    }

    //    internal void Clear()
    //    {
    //        if (TopNode != null && !string.IsNullOrEmpty(TopNode.LocalName) && TopNode.LocalName.ToLower() == "datavalidations")
    //        {
    //            TopNode.ParentNode.RemoveChild(TopNode);
    //        }
    //        _validations.Clear();
    //    }

    //    /// <summary>
    //    /// Adds a validation
    //    /// </summary>
    //    /// <typeparam name="T"></typeparam>
    //    /// <param name="item"></param>
    //    /// <returns></returns>
    //    public IExcelDataValidationWithFormula<T> AddValidation<T>(IExcelDataValidationWithFormula<T> item)
    //        where T : IExcelDataValidationFormula
    //    {
    //        EnsureRootElementExists();
    //        _validations.Add(item);
    //        var formula = item.Formula as ExcelDataValidationFormula;
    //        if(formula != null)
    //        {
    //            formula.RegisterFormulaListener(_formulaListener);
    //        }
    //        OnValidationCountChanged();
    //        return item;
    //    }

    //    /// <summary>
    //    /// Adds a validation
    //    /// </summary>
    //    /// <typeparam name="T"></typeparam>
    //    /// <param name="item"></param>
    //    /// <returns></returns>
    //    public IExcelDataValidationWithFormula2<T> AddValidation<T>(IExcelDataValidationWithFormula2<T> item)
    //        where T : IExcelDataValidationFormula
    //    {
    //        EnsureRootElementExists();
    //        _validations.Add(item);
    //        OnValidationCountChanged();
    //        return item;
    //    }

    //    /// <summary>
    //    /// Returns true if the list is empty
    //    /// </summary>
    //    public bool IsEmpty 
    //    { 
    //        get
    //        {
    //            return _validations == null || _validations.Count == 0;
    //        } 
    //    }

    //    /// <summary>
    //    /// Removes an <see cref="ExcelDataValidation"/> from the collection.
    //    /// </summary>
    //    /// <param name="item">The item to remove</param>
    //    /// <returns>True if remove succeeds, otherwise false</returns>
    //    /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
    //    public bool Remove(IExcelDataValidation item)
    //    {
    //        Require.Argument(item).IsNotNull("item");
    //        if (!(item is ExcelDataValidation))
    //        {
    //            throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
    //        }

    //        ((ExcelDataValidation)item).Delete();
    //        var retVal = _validations.Remove(item);
    //        if (retVal) OnValidationCountChanged();
    //        return retVal;
    //    }

    //    /// <summary>
    //    /// Returns an enumerator that iterates through the collection.
    //    /// </summary>
    //    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    //    public IEnumerator<IExcelDataValidation> GetEnumerator()
    //    {
    //        return _validations.GetEnumerator();
    //    }

    //    /// <summary>
    //    /// Returns an enumerator that iterates through the collection.
    //    /// </summary>
    //    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    //    IEnumerator IEnumerable.GetEnumerator()
    //    {
    //        return _validations.GetEnumerator();
    //    }
    //}
}
