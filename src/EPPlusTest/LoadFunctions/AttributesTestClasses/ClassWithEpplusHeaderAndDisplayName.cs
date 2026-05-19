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
    /// Test class where EpplusTableColumnAttribute.Header IS set.
    /// It should take precedence over DisplayAttribute.Name.
    /// </summary>
    public class ClassWithEpplusHeaderAndDisplayName
    {
        [EpplusTableColumn(Order = 1, Header = "EPPlus Id")]
        [Display(Name = "Display Id")]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 2, Header = "EPPlus Name")]
        [Display(Name = "Display Name")]
        public string Name { get; set; }
    }
}
#endif