using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions.AttributesTestClasses
{
    /// <summary>
    /// Test class where EpplusTableColumnAttribute exists but Order is NOT set.
    /// Should fall back to DisplayAttribute.Order.
    /// </summary>
    public class ClassWithEpplusNoOrderAndDisplayWithOrder
    {
        [EpplusTableColumn(NumberFormat = "0")]
        [Display(Name = "Id Column", Order = 3)]
        public int Id { get; set; }

        [EpplusTableColumn]
        [Display(Name = "Name Column", Order = 1)]
        public string Name { get; set; }

        [EpplusTableColumn]
        [Display(Name = "Description Column", Order = 2)]
        public string Description { get; set; }
    }
}
