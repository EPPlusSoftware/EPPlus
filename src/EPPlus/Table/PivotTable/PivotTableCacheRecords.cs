/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/27/2023         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlusTest.Table.PivotTable
{
    internal class PivotTableCacheRecords 
    {
        internal PivotTableCacheRecords(PivotTableCacheInternal cache)
        {
            Cache = cache;

			var zp = cache._wb._package.ZipPackage;
            if (zp.PartExists(cache.CacheRecordUri))
            {
                var part = zp.GetPart(cache.CacheRecordUri);
                LoadRecords(part);
            }
            else
            {
                CreateRecords();
            }
		}

		private void LoadRecords(ZipPackagePart part)
		{
            var reader = XmlReader.Create(part.GetStream());            
			while (reader.Read())
            {
                if (reader.LocalName=="r" && reader.NodeType==XmlNodeType.Element)
                {
                    AddRecord(reader);
                }
            }
		}
        static CultureInfo _ic = CultureInfo.InvariantCulture;
        static PivotNull _pivotNull = ExcelPivotTable.PivotNullValue;
		private void AddRecord(XmlReader reader)
		{
            int i = 0;
            reader.Read();
			do
            {
                List<object> l;
                if(CacheItems.Count>i)
                {
                    l = CacheItems[i];
                }
                else
                {
                    l = new List<object>();
                    CacheItems.Add(l);
                }
                switch (reader.LocalName)
                {
					case "x":
                        l.Add(int.Parse(reader.GetAttribute("v"), _ic));
                        break;
					case "n":
						l.Add(double.Parse(reader.GetAttribute("v"), _ic));
                        break;
					case "s":
						l.Add(reader.GetAttribute("v"));
                        break;
					case "b":
						l.Add(ConvertUtil.ToBooleanString(reader.GetAttribute("v")));
						break;
					case "d":
						if(ConvertUtil.TryParseDateString(reader.GetAttribute("v"), out DateTime dt))
                        {
                            l.Add(dt);
                        }
                        else
                        {
                            l.Add(_pivotNull);
                        }
						break;
					case "e":
                        var v = reader.GetAttribute("v");
                        if(ExcelErrorValue.IsErrorValue(v))
                        {
							l.Add(ExcelErrorValue.Parse(v));
						}
                        else
                        {
							l.Add(_pivotNull);
						}
						break;                    
                    case "m": 
                        l.Add(_pivotNull);
                        break;
				}
                i++;
			}
            while (reader.Read() && reader.LocalName != "r");
		}

		internal void CreateRecords()
        {
            CacheItems.Clear();
            var sr = Cache.SourceRange;
            var ws = sr.Worksheet;
			var toRow = Cache.GetMaxRow();
			for (int i = 0; i < Cache.Fields.Count; i++)
            {
                var f = Cache.Fields[i];
                Dictionary<object, int> lookup;
				if (f.Grouping?.BaseIndex==null || f.Grouping.BaseIndex == f.Index)
                {
                    lookup = f._cacheLookup;
				}
                else
                {
                    continue;
				}

                var l = new List<object>();
					var c = sr._fromCol + f.Index;

                if (f.IsRowColumnOrPage)
                {
                    f._fieldRecordIndex = new Dictionary<int, List<int>>();
					for (int r = sr._fromRow + 1; r <= toRow; r++) //Skip first row as it contains the headers.
					{
						var v = ws.GetValue(r, c) ?? ExcelPivotTable.PivotNullValue;
						var ix = lookup[v];
						l.Add(ix);
						var ciIx = r - (sr._fromRow + 1);
						if (f._fieldRecordIndex.ContainsKey(ix))
						{
							f._fieldRecordIndex[ix].Add(ciIx);
						}
						else
						{
							f._fieldRecordIndex.Add(ix, new List<int> { ciIx });
						}
					}
				}
                else
                {
                    for (int r = sr._fromRow + 1; r <= toRow; r++)
                    {
                        l.Add(ws.GetValue(r, c) ?? ExcelPivotTable.PivotNullValue);
                    }
                }
                CacheItems.Add(l);
            }
        }
		internal PivotTableCacheInternal Cache { get; }
        public List<List<object>> CacheItems 
        { 
            get; 
        } = new List<List<object>>();
        public int RecordCount
        {
            get
            {
                if(CacheItems==null || CacheItems.Count==0 || CacheItems[0].Count==0)
                {
                    return 0;
                }
                return CacheItems[0].Count;
            }
        }
    }
}
