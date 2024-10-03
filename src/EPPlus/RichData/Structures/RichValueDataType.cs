/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/

internal enum RichValueDataType
{
    /// <summary>
    /// Indicates that the value is a decimal number.
    /// </summary>
    Decimal,
    /// <summary>
    /// Indicates that the value is an Integer
    /// </summary>
    Integer,
    /// <summary>
    ///  Indicates that the value is a Boolean.
    /// </summary>
    Bool,
    /// <summary>
    /// Indicates that the value is an Error. 
    /// </summary>
    Error,
    /// <summary>
    ///  Indicates that the value is a String.
    /// </summary>
    String,
    /// <summary>
    /// Indicates that the value is a Rich Value.
    /// </summary>
    RichValue,
    /// <summary>
    /// Indicates that the value is a Rich Array.
    /// </summary>
    Array,
    /// <summary>
    /// Indicates that the value is a Supporting Property Bag.
    /// </summary>
    SupportingPropertyBag
}
