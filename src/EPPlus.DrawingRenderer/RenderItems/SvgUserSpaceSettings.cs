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
namespace EPPlus.DrawingRenderer
{
    public enum UserSpaceSettings
    {
        /// <summary>
        /// Will use the ObjectBoundingBox as the user space. This is the default for gradients and patterns.
        /// </summary>
        ObjectBoundingBox = 0,
        /// <summary>
        /// Will use the global coordinates as the user space. This is used for gradients and patterns that are outside a group and should be relative to the entire drawing.
        /// </summary>
        UserSpaceOnUse_Global = 1,
        /// <summary>
        /// Will set the user space to the parents coordinates. This is used for gradients and patterns that are inside a group and should be relative to the parent.
        /// </summary>
        UserSpaceOnUse_Parent = 2,
        /// <summary>
        /// Will set the user space to the objects coordinates. This is used for gradients and patterns that are inside a group and should be relative to the object.
        /// </summary>
        UserSpaceOnUse_Object = 3,
    }
}