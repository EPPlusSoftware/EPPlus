/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
#if !NET35
using System.ComponentModel.DataAnnotations;

namespace EPPlusTest.LoadFunctions.AttributesTestClasses
{
    /// <summary>
    /// Test class where only DisplayAttribute is present (no EpplusTableColumnAttribute).
    /// DisplayAttribute.Order should be used in this case.
    /// </summary>
    public class ClassWithDisplayOrderOnly
    {
        [Display(Name = "Id Column", Order = 3)]
        public int Id { get; set; }

        [Display(Name = "Name Column", Order = 1)]
        public string Name { get; set; }

        [Display(Name = "Description Column", Order = 2)]
        public string Description { get; set; }
    }
}
#endif