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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.IO;

namespace OfficeOpenXml
{
    internal class WorksheetZipStream : Stream
    {
        private Stream _stream;
        private long _size;
        private long _bytesRead;
        private int _bufferEnd=0;
        private int _prevBufferEnd = 0;
        public WorksheetZipStream(Stream zip, bool writeToBuffer, long size=-1)
        {
            _stream = zip;
            _size = size;
            _bytesRead = 0;
            WriteToBuffer = writeToBuffer;
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

        byte[] _buffer=null;
        byte[] _prevBuffer, _tempBuffer = new byte[8192];
        public override int Read(byte[] buffer, int offset, int count)
        {
            if(_size>0 && _bytesRead + count > _size)
            {
                count = (int)(_size - _bytesRead);
            }
            if (_buffer != null)
            {
                if(_tempBuffer.Length<_bufferEnd) _tempBuffer = new byte[_bufferEnd];
                Array.Copy(_buffer, _tempBuffer, _bufferEnd);
            }

            var r =_stream.Read(buffer, offset, count);
            if(r>0)
            {
                _prevBuffer = _tempBuffer;
                _prevBufferEnd = _bufferEnd;

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
        public BinaryWriter Buffer=new BinaryWriter(RecyclableMemory.GetStream());
        public void SetWriteToBuffer()
        {
            Buffer = new BinaryWriter(RecyclableMemory.GetStream());
            if (WriteToBuffer==false)
            {
                if (_prevBuffer != null)
                {
                    Buffer.Write(_prevBuffer, 0, _prevBufferEnd);
                }
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
        internal string GetBufferAsStringRemovingElement(bool writeToBufferAfter, string element)
        {
            WriteToBuffer = writeToBufferAfter;
            Buffer.Flush();
            var xml = System.Text.Encoding.UTF8.GetString(((MemoryStream)Buffer.BaseStream).ToArray());            
            GetElementPos(xml, element, out int startIx, out int endIx);
            if(startIx > 0)
            {                
                return xml.Substring(0, startIx) + GetPlaceholderTag(xml, startIx, endIx);
            }
            else
            {
                return xml;
            }
        }

        private static string GetPlaceholderTag(string xml, int startIx, int endIx)
        {
            var placeholderTag = xml.Substring(startIx, endIx - startIx);
            placeholderTag = placeholderTag.Replace("/", "");
            placeholderTag = placeholderTag.Substring(0, placeholderTag.Length - 1) + "/>";
            return placeholderTag;
        }

        private int GetEndElementPos(string xml, string element, int endIx)
        {
            var ix = xml.IndexOf("/" + element + ">", endIx);
            if(ix > 0)
            {
                return ix + element.Length + 2;
            }
            return -1;
        }

        private void GetElementPos(string xml, string element, out int startIx, out int endIx)
        {
            int ix = -1;
            do
            {
                ix = xml.IndexOf(element, ix + 1);
                if (ix > 0 && (xml[ix - 1] == ':' || xml[ix - 1] == '<'))
                {
                    startIx = ix;
                    if (startIx >= 0 && xml[startIx] != '<')
                    {
                        startIx--;
                    }
                    endIx = ix + element.Length;
                    while (endIx < xml.Length && xml[endIx] == ' ')
                    {
                        endIx++;
                    }
                    if(endIx<xml.Length && xml[endIx] == '>')                       
                    {
                        endIx++;
                        return;
                    }
                    else if(endIx < xml.Length + 1 && xml.Substring(endIx, 2) == "/>")
                    {
                        endIx += 2;
                        return;
                    }
                }
            }
            while (ix >= 0);
            startIx = endIx = -1;
        }

        internal void ReadToEnd()
        {
            if(_bytesRead < _size)
            {                
                var r = _stream.Read(_buffer, 0, (int)(_size - _bytesRead));
                if (WriteToBuffer == false)
                {
                    Buffer.Write(_buffer, 0, _bufferEnd);
                }
                _bytesRead = _size;
            }
        }

        internal string ReadFromEndElement(string endElement, string startXml="", string readToElement=null, bool writeToBuffer = true, string xmlPrefix = "", string attribute="")
        {
            if(string.IsNullOrEmpty(readToElement) && _stream.Position < _stream.Length)
            {
                ReadToEnd();
            }
            
            Buffer.Flush();
            var xml = System.Text.Encoding.UTF8.GetString(((MemoryStream)Buffer.BaseStream).ToArray());
            var endElementIx = FindElementPos(xml, endElement, false);

            if (endElementIx < 0) return startXml;
            if(string.IsNullOrEmpty(readToElement))
            {
                xml = xml.Substring(endElementIx);
            }
            else
            {
                var toElementIx = FindElementPos(xml, readToElement);
                if(toElementIx>=endElementIx)
                {
                    xml =  xml.Substring(endElementIx, toElementIx - endElementIx);
                    xml += string.IsNullOrEmpty(xmlPrefix) ? $"<{readToElement}{attribute}/>" : $"<{xmlPrefix}:{readToElement}{attribute}/>";
                }
                else
                {
                    xml = xml.Substring(endElementIx);
                }
            }
            WriteToBuffer = writeToBuffer;
            return startXml + xml;
        }

        /// <summary>
        /// Returns the position in the xml document for an element. Either returns the position of the start element or the end element.
        /// </summary>
        /// <param name="xml">The xml to search</param>
        /// <param name="element">The element</param>
        /// <param name="returnStartPos">If the position before the start element is returned. If false the end of the end element is returned.</param>
        /// <returns>The position of the element in the input xml</returns>
        private int FindElementPos(string xml, string element, bool returnStartPos = true)
        {
            var ix = 0;
            while (true)
            {
                ix = xml.IndexOf(element, ix);
                if (ix > 0 && ix < xml.Length-1)
                {                    
                    var c = xml[ix + element.Length];
                    if (c == '>' || c == ' ' || c == '/')
                    {
                        c = xml[ix - 1];
                        if (c != '/' && c != ':' && xml[ix - 1] != '<') return -1;

                        if (returnStartPos)
                        {
                            return xml.LastIndexOf('<', ix);
                        }
                        else
                        {
                            //Return the end element, either </element> or <element/>
                            var startIx = xml.LastIndexOf("<", ix);
                            if (ix > 0)
                            {
                                var end = xml.IndexOf(">", ix + element.Length - 1);
                                if (xml[startIx + 1] == '/' || xml[end - 1] == '/')
                                {
                                    return end + 1;
                                }
                            }
                        }
                    }
                }
                if(ix<=0) return -1;
                ix+=element.Length;
            }
        }
    }
}