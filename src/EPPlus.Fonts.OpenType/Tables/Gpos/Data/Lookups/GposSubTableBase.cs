/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS subtable base class
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// Base class for all GPOS subtables.
    /// Each lookup type has one or more subtable formats that inherit from this.
    /// </summary>
    public abstract class GposSubTableBase : FontTableElement
    {
        // Currently empty - provides type safety and future extensibility
        // All GPOS subtables inherit from this for polymorphic storage in Lookup.SubTables
    }
}