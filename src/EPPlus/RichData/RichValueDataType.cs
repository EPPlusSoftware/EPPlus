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
