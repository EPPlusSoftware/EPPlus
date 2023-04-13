using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class SortByComparerTests
    {
        [TestMethod]
        public void ShouldSortNullValues_Asc()
        {
            var comparer = new SortByComparer();

            var list = new List<object>
            { 3, 2, null, 4, 5 };

            list.Sort((a, b) => comparer.Compare(a, b, 1));

            Assert.AreEqual(2, list[0]);
            Assert.IsNull(list.Last());
        }

        [TestMethod]
        public void ShouldSortNullValues_Desc()
        {
            var comparer = new SortByComparer();

            var list = new List<object>
            { 3, 2, null, 4, 5 };

            list.Sort((a, b) => comparer.Compare(a, b, -1));

            Assert.AreEqual(5, list[0]);
            Assert.IsNull(list.Last());
        }
    }
}
