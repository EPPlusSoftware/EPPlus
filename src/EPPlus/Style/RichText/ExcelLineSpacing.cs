using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Line spacing
    /// </summary>
    public enum eLineSpacing
    {
        /// <summary>
        /// Single line spacing
        /// </summary>
        Single,
        /// <summary>
        /// 1.5 lines
        /// </summary>
        OneAndAHalf,
        /// <summary>
        /// Double line spacing
        /// </summary>
        Double,
        /// <summary>
        /// Exact point spacing
        /// </summary>
        Exactly,
        /// <summary>
        /// Multiple line spacing
        /// </summary>
        Multiple
    }
    
    ///// <summary>
    ///// Used to define line-spacing in pargraphs/shapes
    ///// Default is single line spacing
    ///// </summary>
    //public class ExcelLineSpacing
    //{
    //    //public double Before;
    //    //public double After;

    //    eLineSpacing _lineSpacingType;

    //    /// <summary>
    //    /// Only used when setting Exactly or Multiple
    //    /// </summary>
    //    public double ptOrPercentValue;


    //    /// <summary>
    //    /// If setting Exactly or Multiple it is recommended to use
    //    /// SetExactly or SetMultiple functions.
    //    /// Otherwise they are set to default values 13.2 or 3
    //    /// </summary>
    //    public eLineSpacing LineSpacingType 
    //    { 
    //        get 
    //        { 
    //            return _lineSpacingType; 
    //        } 
    //        set 
    //        {
    //            if (_lineSpacingType == eLineSpacing.Exactly)
    //            {
    //                SetExactly(13.2);
    //            }
    //            else if (_lineSpacingType == eLineSpacing.Multiple)
    //            {
    //                SetMultiple(3);
    //            }
    //            else
    //            {
    //                ptOrPercentValue = 0;
    //                _lineSpacingType = value;
    //            }
    //        } 
    //    }

    //    internal ExcelLineSpacing()
    //    {
    //        LineSpacingType = eLineSpacing.Single;
    //    }
    //    /// <summary>
    //    /// Sets line spacing to Exactly inPoints
    //    /// </summary>
    //    /// <param name="inPoints"></param>
    //    public void SetExactly(double inPoints)
    //    {
    //        _lineSpacingType = eLineSpacing.Exactly;
    //        ptOrPercentValue = inPoints;
    //    }
    //    /// <summary>
    //    /// Sets line spacing to multiple of percent
    //    /// </summary>
    //    /// <param name="percent"></param>
    //    public void SetMultiple(double percent)
    //    {
    //        _lineSpacingType = eLineSpacing.Multiple;
    //        ptOrPercentValue = percent;
    //    }

    //    ///// <summary>
    //    ///// If setting Exactly or Multiple it is recommended to use
    //    ///// SetExactly or SetMultiple functions.
    //    ///// Otherwise they are set to default values 13.2 or 3
    //    ///// </summary>
    //    ///// <param name="lineSpacingType"></param>
    //    //public void SetSpacingType(eLineSpacing lineSpacingType)
    //    //{
    //    //    LineSpacingType = lineSpacingType;
    //    //    if(lineSpacingType == eLineSpacing.Exactly)
    //    //    {
    //    //        SetExactly(13.2);
    //    //    }
    //    //    else if(lineSpacingType == eLineSpacing.Multiple)
    //    //    {
    //    //        SetMultiple(3);
    //    //    }
    //    //}
    //}
}
