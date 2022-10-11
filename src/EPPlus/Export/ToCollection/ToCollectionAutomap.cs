using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
namespace OfficeOpenXml.Export.ToCollection
{
#if(!NET35)
    internal class ToCollectionAutomap
    {
        internal static List<Tuple<int, PropertyInfo>> GetAutomapList<T>(List<string> h)
        {
            var t = typeof(T);

            var pl = new List<Tuple<int, PropertyInfo>>();
            foreach (var m in t.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                var ix = h.FindIndex(x => RemoveWS(x).Equals(m.Name, StringComparison.CurrentCultureIgnoreCase));
                if (ix < 0)
                {
                    var tca = m.GetFirstAttributeOfType<EpplusTableColumnAttributeBase>();
                    if (tca != null)
                    {
                        ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(tca.Header), StringComparison.CurrentCultureIgnoreCase));
                    }
                    if (ix < 0)
                    {
                        var da = m.GetFirstAttributeOfType<DescriptionAttribute>();
                        if (da != null)
                        {
                            ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(da.Description), StringComparison.CurrentCultureIgnoreCase));
                        }
                        if (ix < 0)
                        {
                            var dna = m.GetFirstAttributeOfType<DisplayNameAttribute>();
                            if (dna != null)
                            {
                                ix = h.FindIndex(x => RemoveWS(x).Equals(RemoveWS(dna.DisplayName), StringComparison.CurrentCultureIgnoreCase));
                            }
                        }
                    }
                }
                if (ix >= 0)
                {
                    pl.Add(new Tuple<int, PropertyInfo>(ix, m));
                }
            }
            return pl;
        }

        private static string RemoveWS(string v)
        {
            return v.Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n","");
        }
    }
#endif
}
