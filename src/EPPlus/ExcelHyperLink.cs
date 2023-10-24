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
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// HyperlinkClass
    /// </summary>
    public class ExcelHyperLink : Uri
    {
        /// <summary>
        /// A new hyperlink with the specified URI
        /// </summary>
        /// <param name="uriString">The URI</param>
        public ExcelHyperLink(string uriString) :
            base(uriString)
        {
            OriginalUri = this;
        }
#if !Core
        /// <summary>
        /// A new hyperlink with the specified URI. This syntax is obsolete
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="dontEscape"></param>
        [Obsolete("base constructor 'System.Uri.Uri(string, bool)' is obsolete: 'The constructor has been deprecated. Please use new ExcelHyperLink(string). The dontEscape parameter is deprecated and is always false.")]
        public ExcelHyperLink(string uriString, bool dontEscape) :
            base(uriString, dontEscape)
        {
            OriginalUri = (Uri)this;
        }
#endif
        /// <summary>
        /// A new hyperlink with the specified URI and kind
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="uriKind">Kind (absolute/relative or indeterminate)</param>
        public ExcelHyperLink(string uriString, UriKind uriKind) :
            base(uriString, uriKind)
        {
            OriginalUri = this;
        }
        /// <summary>
        /// Sheet internal reference
        /// </summary>
        /// <param name="referenceAddress">The address or defined name</param>
        /// <param name="display">Displayed text</param>
        public ExcelHyperLink(string referenceAddress, string display) :
            base("xl://internal")   //URI is not used on internal links so put a dummy uri here.
        {
            _referenceAddress = referenceAddress;
            _display = display;
        }
        
        string _referenceAddress = null;
        /// <summary>
        /// The Excel address for internal links or extended data for external hyper links not supported by the Uri class.
        /// </summary>
        public string ReferenceAddress
        {
            get
            {
                return _referenceAddress;
            }
            set
            {
                _referenceAddress = value;
            }
        }
        string _display = "";
        /// <summary>
        /// Displayed text
        /// </summary>
        public string Display
        {
            get
            {
                return _display;
            }
            set
            {
                _display = value;
            }
        }
        /// <summary>
        /// Tooltip
        /// </summary>
        public string ToolTip
        {
            get;
            set;
        }
        int _colSpan = 0;
        /// <summary>
        /// If the hyperlink spans multiple columns
        /// </summary>
        public int ColSpan
        {
            get
            {
                return _colSpan;
            }
            set
            {
                _colSpan = value;
            }
        }
        int _rowSpan = 0;
        /// <summary>
        /// If the hyperlink spans multiple rows
        /// </summary>
        public int RowSpan
        {
            get
            {
                return _rowSpan;
            }
            set
            {
                _rowSpan = value;
            }
        }
        /// <summary>
        /// Used to handle non absolute URI's. 
        /// Is used if IsAblsoluteUri is true. The base URI will have a dummy value of xl://nonAbsolute.
        /// </summary>
        public Uri OriginalUri
        {
            get;
            internal set;
        }
        internal string RId
        {
            get;
            set;
        }
        internal string Target { get; set; }
    }
}
