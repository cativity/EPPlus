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

namespace EPPlusTest.Utils;

/// <summary>
/// A buffer that rolls out memory as it's written to the buffer. 
/// </summary>
internal class RollingBuffer
{
    bool _isRolling;
    byte[] _buffer;
    int _index;

    internal RollingBuffer(int size)
    {
        this._buffer = new byte[size];
    }

    internal void Write(byte[] bytes, int size = -1)
    {
        if (size < 0)
        {
            size = bytes.Length;
        }

        if (size >= this._buffer.Length)
        {
            this._index = 0;
            this._isRolling = true;
            Array.Copy(bytes, size - this._buffer.Length, this._buffer, 0, this._buffer.Length);
        }
        else if (size + this._index > this._buffer.Length)
        {
            int endSize = this._buffer.Length - this._index;
            this._isRolling = true;

            if (endSize > 0)
            {
                Array.Copy(bytes, 0, this._buffer, this._index, endSize);
            }

            this._index = size - endSize;
            Array.Copy(bytes, endSize, this._buffer, 0, this._index);
        }
        else
        {
            Array.Copy(bytes, 0, this._buffer, this._index, size);
            this._index += size;
        }
    }

    internal byte[] GetBuffer()
    {
        byte[] ret;

        if (this._isRolling)
        {
            ret = new byte[this._buffer.Length];
            Array.Copy(this._buffer, this._index, ret, 0, this._buffer.Length - this._index);

            if (this._index > 0)
            {
                Array.Copy(this._buffer, 0, ret, this._buffer.Length - this._index, this._index);
            }
        }
        else
        {
            ret = new byte[this._index];
            Array.Copy(this._buffer, ret, ret.Length);
        }

        return ret;
    }
}