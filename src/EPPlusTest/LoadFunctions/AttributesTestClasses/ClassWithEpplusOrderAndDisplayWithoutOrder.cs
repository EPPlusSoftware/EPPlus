/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
#if !NET35
using System.ComponentModel.DataAnnotations;

namespace EPPlusTest.LoadFunctions.AttributesTestClasses
{
    /// <summary>
    /// Test class where EpplusTableColumnAttribute has Order set but
    /// DisplayAttribute does NOT have Order set.
    /// EpplusTableColumnAttribute.Order should be used.
    /// </summary>
    public class ClassWithEpplusOrderAndDisplayWithoutOrder
    {
        [EpplusTableColumn(Order = 3)]
        [Display(Name = "Id Column")]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 1)]
        [Display(Name = "Name Column")]
        public string Name { get; set; }

        [EpplusTableColumn(Order = 2)]
        [Display(Name = "Description Column")]
        public string Description { get; set; }
    }
}
#endif