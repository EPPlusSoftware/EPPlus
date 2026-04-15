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
    /// Test class using DisplayAttribute with ResourceType.
    /// When ResourceType is set, GetName() should be used instead of Name
    /// to get the localized value from the resource class.
    /// </summary>
    public class ClassWithDisplayResourceType
    {
        [EpplusTableColumn(Order = 1)]
        [Display(Name = "IdHeader", ResourceType = typeof(LoadFromCollResources))]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 2)]
        [Display(Name = "NameHeader", ResourceType = typeof(LoadFromCollResources))]
        public string Name { get; set; }

        [EpplusTableColumn(Order = 3)]
        [Display(Name = "DescriptionHeader", ResourceType = typeof(LoadFromCollResources))]
        public string Description { get; set; }
    }
}
#endif