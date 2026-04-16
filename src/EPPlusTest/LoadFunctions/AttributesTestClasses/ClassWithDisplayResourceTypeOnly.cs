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
    /// Test class using DisplayAttribute with ResourceType but without EpplusTableColumnAttribute.
    /// GetName() should return the localized value from the resource class.
    /// </summary>
    public class ClassWithDisplayResourceTypeOnly
    {
        [Display(Name = "IdHeader", ResourceType = typeof(LoadFromCollResources), Order = 1)]
        public int Id { get; set; }

        [Display(Name = "NameHeader", ResourceType = typeof(LoadFromCollResources), Order = 2)]
        public string Name { get; set; }

        [Display(Name = "DescriptionHeader", ResourceType = typeof(LoadFromCollResources), Order = 3)]
        public string Description { get; set; }
    }
}
#endif