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
    /// Test class where EpplusTableColumnAttribute has Order set but NOT Header.
    /// Header should fall back to DisplayAttribute.GetName().
    /// </summary>
    public class ClassWithEpplusOrderAndDisplayName
    {
        [EpplusTableColumn(Order = 3)]
        [Display(Name = "The Id")]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 1)]
        [Display(Name = "The Name")]
        public string Name { get; set; }

        [EpplusTableColumn(Order = 2)]
        [Display(Name = "The Description")]
        public string Description { get; set; }
    }
}
#endif