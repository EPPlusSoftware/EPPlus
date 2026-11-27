/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
namespace EPPlus.DrawingRenderer
{
    public abstract class BaseRenderer<T, T2>
    {
        protected BaseRenderer(T outputStream)
        {
            OutputStream = outputStream;
        }
        public T OutputStream { get; }
        public abstract void Render(T2 item);
    }
}
