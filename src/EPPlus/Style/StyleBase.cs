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
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Base class for styles
    /// </summary>
    public abstract class StyleBase
    {
        internal ExcelStyles _styles;
        internal OfficeOpenXml.XmlHelper.ChangedEventHandler _ChangedEvent;
        internal int _positionID;
        internal string _address;
        internal StyleBase(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address)
        {
            _styles = styles;
            _ChangedEvent = ChangedEvent;
            _address = Address;
            _positionID = PositionID;
        }
        internal int Index { get; set;}
        internal abstract string Id {get;}

        internal virtual void SetIndex(int index)
        {
            Index = index;
        }
    }
}
