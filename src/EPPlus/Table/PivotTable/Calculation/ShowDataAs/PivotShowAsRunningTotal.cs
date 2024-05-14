/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 01/18/2024         EPPlus Software AB       EPPlus 7.2
*************************************************************************************************/
using EPPlusTest.Table.PivotTable;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs
{
    internal class PivotShowAsRunningTotal : PivotShowAsBase
    {
		protected static int _bf;
        protected static int _colFieldsStart;
        protected static int _keyCol;
        protected static PivotTableCacheRecords _record;
        protected static int _maxBfKey;

        internal override void Calculate(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems)
		{
			CalculateRunningTotal(df, fieldIndex, ref calculatedItems, false);
		}

		internal void CalculateRunningTotal(ExcelPivotTableDataField df, List<int> fieldIndex, ref PivotCalculationStore calculatedItems, bool leaveParentSum)
		{
			_bf = fieldIndex.IndexOf(df.BaseField);

            if (_bf < 0)
            {
                foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
                {
                    calculatedItems[key.Key] = ErrorValues.NAError;
                }
				return;
            }

            _colFieldsStart = df.Field.PivotTable.RowFields.Count;
			_keyCol = fieldIndex.IndexOf(df.BaseField);
			_record = df.Field.PivotTable.CacheDefinition._cacheReference.Records;
			_maxBfKey = 0;

			if (_record.CacheItems[df.BaseField].Count(x => x is int) > 0)
			{
				_maxBfKey = (int)_record.CacheItems[df.BaseField].Where(x => x is int).Max();
			}

            foreach (var key in calculatedItems.Index.OrderBy(x => x.Key, ArrayComparer.Instance))
            {
                if (IsSumBefore(key.Key, _bf, fieldIndex, _colFieldsStart))
                {
                    if (!(leaveParentSum == true && key.Key[_keyCol] == PivotCalculationStore.SumLevelValue))
                    {
                        calculatedItems[key.Key] = null;
                    }
                }
                else if (IsSumAfter(key.Key, _bf, fieldIndex, _colFieldsStart) == true)
                {
                    if (key.Key[_keyCol] > 0)
                    {
                        var prevKey = GetPrevKey(key.Key, _keyCol);
                        if (calculatedItems.ContainsKey(prevKey))
                        {
                            if (calculatedItems[key.Key] is double current)
                            {
                                if (calculatedItems[prevKey] is double prev)
                                {
                                    calculatedItems[key.Key] = current + prev;
                                }
                                else
                                {
                                    calculatedItems[key.Key] = calculatedItems[prevKey]; //The prev key is an error, set the value to that error.
                                }
                            }
                        }
                    }

                    if (key.Key[_keyCol] < _maxBfKey)
                    {
                        var nextKey = GetNextKey(key.Key, _keyCol);
                        while (nextKey[_keyCol] < _maxBfKey && calculatedItems.ContainsKey(nextKey) == false)
                        {
                            calculatedItems[nextKey] = calculatedItems[key.Key];
                            nextKey = GetNextKey(nextKey, _keyCol);
                        }
                    }
                }
            }
        }
	}
}
