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
    /// Test class with a negative Order value on EpplusTableColumnAttribute,
    /// matching the customer's scenario (Order = -90).
    /// </summary>
    public class ClassWithNegativeEpplusOrder
    {
        [EpplusTableColumn(Order = -90)]
        [Display(Name = "NumRegistro", Order = 5)]
        public int? NumRegistro { get; set; }

        [EpplusTableColumn(Order = 1)]
        [Display(Name = "Nombre", Order = 1)]
        public string Nombre { get; set; }

        [EpplusTableColumn(Order = 2)]
        [Display(Name = "Descripcion", Order = 2)]
        public string Descripcion { get; set; }
    }
}
#endif