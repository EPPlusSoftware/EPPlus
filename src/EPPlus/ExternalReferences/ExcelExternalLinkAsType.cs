/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/28/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Provides a simple way to type cast <see cref="ExcelExternalLink"/> object top its top level class.
    /// </summary>
    public class ExcelExternalLinkAsType
    {
        ExcelExternalLink _externalLink;
        internal ExcelExternalLinkAsType(ExcelExternalLink externalLink)
        {
            _externalLink = externalLink;
        }
        /// <summary>
        /// Converts the external link to it's top level .
        /// </summary>
        /// <typeparam name="T">The type of external link. T must be inherited from ExcelExternalLink</typeparam>
        /// <returns>The external link as type T</returns>
        public T Type<T>() where T : ExcelExternalLink
        {
            return _externalLink as T;
        }
        /// <summary>
        /// Return the external link as an external workbook. If the external link is not of type <see cref="ExcelExternalBook" />, null is returned
        /// </summary>
        public ExcelExternalWorkbook ExternalWorkbook
        {
            get
            {
                return _externalLink as ExcelExternalWorkbook;
            }
        }
        /// <summary>
        /// Return the external link as a dde link. If the external link is not of type <see cref="ExcelExternalDdeLink"/>, null is returned
        /// </summary>
        public ExcelExternalDdeLink DdeLink
        {
            get
            {
                return _externalLink as ExcelExternalDdeLink;
            }
        }
        /// <summary>
        /// Return the external link as a ole link. If the external link is not of type <see cref="ExcelExternalOleLink"/>, null is returned
        /// </summary>
        public ExcelExternalOleLink OleLink
        {
            get
            {
                return _externalLink as ExcelExternalOleLink;
            }
        }

    }
}