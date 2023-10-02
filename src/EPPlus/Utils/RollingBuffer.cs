/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/06/2023         EPPlus Software AB           Added
 *************************************************************************************************/
using System;
namespace EPPlusTest.Utils
{
    /// <summary>
    /// A buffer that rolls out memory as it's written to the buffer. 
    /// </summary>
    internal class RollingBuffer
    {
        bool _isRolling = false;
        byte[] _buffer;
        int _index = 0;
        internal RollingBuffer(int size)
        {
            _buffer= new byte[size];                 
        }
        internal void Write(byte[] bytes, int size=-1)
        {
            if (size < 0) size = bytes.Length;
            if(size >= _buffer.Length)
            {
                _index = 0;
                _isRolling = true;
                Array.Copy(bytes, size - _buffer.Length, _buffer, 0, _buffer.Length);
            }
            else if(size + _index > _buffer.Length)
            {
                var endSize = _buffer.Length - _index;
                _isRolling = true;
                if(endSize > 0)
                {
                    Array.Copy(bytes, 0, _buffer, _index, endSize);
                }
                _index = size - endSize;
                Array.Copy(bytes, endSize, _buffer, 0, _index);
            }
            else
            {
                Array.Copy(bytes, 0, _buffer, _index, size);
                _index += size;
            }
        }
        internal byte[] GetBuffer()
        {
            byte[] ret;
            if(_isRolling)
            {
                ret = new byte[_buffer.Length];
                Array.Copy(_buffer, _index, ret,0,_buffer.Length-_index);
                if(_index>0)
                {
                    Array.Copy(_buffer, 0, ret, _buffer.Length - _index, _index);
                }
            }
            else
            {
                ret = new byte[_index];
                Array.Copy(_buffer, ret, ret.Length);
            }
            return ret;
        }
    }
}
