/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.Constants
{
    internal static class StructureKeyNames
    {
        internal static class Errors
        {
            internal static class  FieldError
            {
                public const string ErrorType = "errorType";
                public const string Field = "field";
            }

            internal static class PropagatedError
            {
                public const string ErrorType = "errorType";
                public const string Propagated = "propagated";
            }

            internal static class Spill
            {
                public const string ColOffset = "colOffset";
                public const string ErrorType = "errorType";
                public const string RwOffset = "rwOffset";
                public const string SubType = "subType";
            }

            internal static class WithSubType
            {
                public const string ErrorType = "errorType";
                public const string SubType = "subType";
            }
        }

        internal static class LocalImages
        {
            internal static class Image
            {
                public const string RelLocalImageIdentifier = "_rvRel:LocalImageIdentifier";
                public const string CalcOrigin = "CalcOrigin";
            }

            internal static class ImageAltText
            {
                public const string RelLocalImageIdentifier = "_rvRel:LocalImageIdentifier";
                public const string CalcOrigin = "CalcOrigin";
                public const string Text = "Text";
            }
        }
    }
}
