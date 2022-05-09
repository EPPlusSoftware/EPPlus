/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Packaging.Ionic.Zip;
using System.IO;

namespace OfficeOpenXml
{
    internal class WorksheetZipStream : Stream
    {
        private Stream _stream;
        private long _size;
        private long _bytesRead;
        private int _bufferEnd=0;
        public WorksheetZipStream(Stream zip, bool writeToBuffer, long size=-1)
        {
            _stream = zip;
            _size = size;
            _bytesRead = 0;
            WriteToBuffer = writeToBuffer; ;
        }

        public override bool CanRead => _stream.CanRead;

        public override bool CanSeek => _stream.CanSeek;

        public override bool CanWrite => _stream.CanWrite;

        public override long Length => _stream.Length;

        public override long Position { get => _stream.Position; set => _stream.Position=value; }

        public override void Flush()
        {
            _stream.Flush();
        }

        byte[] _buffer; 
        public override int Read(byte[] buffer, int offset, int count)
        {
            if(_size>0 && _bytesRead + count > _size)
            {
                count = (int)(_size - _bytesRead);
            }
            var r=_stream.Read(buffer, offset, count);
            if(r>0)
            {
                _buffer = buffer;
                _bytesRead += r;
                _bufferEnd = r;

                if (WriteToBuffer)
                {
                    Buffer.Write(buffer, 0, r);
                }

            }
            return r;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _stream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            _stream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _stream.Write(buffer, offset, count);
        }
        public BinaryWriter Buffer=new BinaryWriter(new MemoryStream());
        public void SetWriteToBuffer()
        {
            Buffer = new BinaryWriter(new MemoryStream());
            if (WriteToBuffer==false)
            {
                Buffer.Write(_buffer,0, _bufferEnd);
            }

            WriteToBuffer = true;
        }
        public bool WriteToBuffer { get; set; }

        internal string GetBufferAsString(bool writeToBufferAfter)
        {
            WriteToBuffer = writeToBufferAfter;
            Buffer.Flush();
            return System.Text.Encoding.UTF8.GetString(((MemoryStream)Buffer.BaseStream).ToArray());
        }

        internal void ReadToEnd()
        {
            if(_bytesRead<_size)
            {
                
                var r = _stream.Read(_buffer, 0, (int)(_size-_bytesRead));
                if (WriteToBuffer == false)
                {
                    Buffer.Write(_buffer, 0, _bufferEnd);
                }
                _bytesRead = _size;
            }
        }
    }
}