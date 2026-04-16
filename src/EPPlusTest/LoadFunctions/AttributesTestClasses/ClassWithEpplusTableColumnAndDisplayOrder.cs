/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35
using System.ComponentModel.DataAnnotations;
#endif

namespace EPPlusTest.LoadFunctions.AttributesTestClasses
{
#if !NET35
    /// <summary>
    /// Test class where EpplusTableColumnAttribute.Order should take precedence
    /// over DisplayAttribute.Order when both are present on the same property.
    /// </summary>
    public class ClassWithEpplusTableColumnAndDisplayOrder
    {
        // EpplusTableColumn Order = 3, Display Order = 1
        // Expected column order should be 3 (EpplusTableColumn wins)
        [EpplusTableColumn(Order = 3)]
        [Display(Name = "Id Column", Order = 1)]
        public int Id { get; set; }

        // EpplusTableColumn Order = 1, Display Order = 3
        // Expected column order should be 1 (EpplusTableColumn wins)
        [EpplusTableColumn(Order = 1)]
        [Display(Name = "Name Column", Order = 3)]
        public string Name { get; set; }

        // EpplusTableColumn Order = 2, Display Order = 2
        [EpplusTableColumn(Order = 2)]
        [Display(Name = "Description Column", Order = 2)]
        public string Description { get; set; }
    }
#endif
}