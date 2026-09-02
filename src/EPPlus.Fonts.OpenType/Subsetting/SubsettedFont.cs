/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    public sealed class SubsettedFont
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="family">canonical, pre-subset family name</param>
        /// <param name="subFamily"></param>
        /// <param name="font">the subsetted instance to embed</param>
        internal SubsettedFont(string family, FontSubFamily subFamily, OpenTypeFont font)
        {
            Family = family; SubFamily = subFamily; Font = font;
        }

        /// <summary>
        /// canonical, pre-subset family name
        /// </summary>
        public string Family { get; }
        public FontSubFamily SubFamily { get; }
        /// <summary>
        /// the subsetted instance to embed
        /// </summary>
        public OpenTypeFont Font { get; }
    }
}
