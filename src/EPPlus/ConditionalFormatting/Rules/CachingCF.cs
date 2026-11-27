using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class CachingCF : ExcelConditionalFormattingRule
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        /// <param name="cfType"></param>
        internal CachingCF(
          eExcelConditionalFormattingRuleType cfType,
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(cfType, address, priority, worksheet)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        /// <param name="cfType"></param>
        internal CachingCF(eExcelConditionalFormattingRuleType cfType, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(cfType, address, ws, xr)
        {
        }

        internal CachingCF(CachingCF copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new CachingCF(this, newWs);
        }

        protected List<object> cellValueCache = new List<object>();
        /// <summary>
        /// True when <see cref="cellValueCache"/> has been filled for the current export.
        /// </summary>
        protected bool _cellValueCacheUpdated = false;

        protected virtual void UpdateCellValueCache(bool asStrings = false, bool cacheOnlyNumeric = false)
        {
            cellValueCache.Clear();

            //Only visit cells that exist. The rule's range can cover a full sheet.
            foreach (var value in CFRangeValues.GetValues(_ws, Address))
            {
                if (asStrings)
                {
                    cellValueCache.Add(value.ToString());
                }
                else
                {
                    if (cacheOnlyNumeric)
                    {
                        if (value.IsNumeric())
                        {
                            cellValueCache.Add(value);
                        }
                    }
                    else
                    {
                        cellValueCache.Add(value);
                    }
                }
            }
            _cellValueCacheUpdated = true;
        }

        internal override void RemoveTempExportData()
        {
            base.RemoveTempExportData();
            cellValueCache.Clear();
            _cellValueCacheUpdated = false;
        }
    }
}
