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
using System.Text;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// A collection of the module level attributes
    /// </summary>
    public class ExcelVbaModuleAttributesCollection : ExcelVBACollectionBase<ExcelVbaModuleAttribute>
    {
        internal string GetAttributeText()
        {
            StringBuilder sb=new StringBuilder();

            foreach (var attr in this)
            {
                sb.AppendFormat("Attribute {0} = {1}\r\n", attr.Name, attr.DataType==eAttributeDataType.String ? "\"" + attr.Value + "\"" : attr.Value);
            }
            return sb.ToString();
        }
    }
}