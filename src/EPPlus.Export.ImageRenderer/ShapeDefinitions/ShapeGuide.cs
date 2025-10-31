/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Diagnostics;

namespace EPPlusImageRenderer.ShapeDefinitions
{
    /// <summary>
    /// 20.1.9.11 Ecma part 1
    /// 20.1.10.76 - Preset Text Shape Types
    /// </summary>
    [DebuggerDisplay("{Name}-{Formula}={CalculatedValue}")]
    public class ShapeGuide
    {
        public string Name { get; set; }
        public string Formula { get; set; }
        public double CalculatedValue 
        { 
            get; 
            set; 
        }

        internal ShapeGuide Clone()
        {
            return new ShapeGuide() { Name=Name, Formula=Formula, CalculatedValue=CalculatedValue};
        }
    }
}
