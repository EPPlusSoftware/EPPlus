using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Types
{
    /// <summary>
    /// Flags used for rich data.
    /// </summary>
    [Flags]
    internal enum RichValueKeyFlags
    {
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) in the default Card View
        /// </summary>
        ShowInCardView,
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) from formulas and the object model
        /// </summary>
        ShowInDotNotation,
        /// <summary>
        /// False indicates that we hide this key value pair (KVP) from AutoComplete, sort, filter, and Find
        /// </summary>
        ShowInAutoComplete,
        /// <summary>
        /// True indicates that we do not write this key value pair (KVP) into the file, it only exists in memory
        /// </summary>
        ExcludeFromFile,
        /// <summary>
        /// True indicates that we exclude this key value pair (KVP) when comparing rich values.
        /// </summary>
        ExcludeFromCalcComparison,
    }
}
