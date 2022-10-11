/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/04/2022         EPPlus Software AB       Initial release EPPlus 6.1
 *************************************************************************************************/
namespace OfficeOpenXml.Export.ToCollection
{
    /// <summary>
    /// How conversion failures should be handled when mapping properties in the ToCollection method.
    /// </summary>
    public enum ToCollectionConversionFailureStrategy
    {
        /// <summary>
        /// Throw an Exception if the conversion fails. Blank values will return the default value for the type. An <see cref="Exceptions.EPPlusDataTypeConvertionException"/> will be thrown on any datatype conversion failure when mapping properties.
        /// </summary>
        Exception,
        /// <summary>
        /// Set the default value for the property.
        /// </summary>
        SetDefaultValue
    }
}