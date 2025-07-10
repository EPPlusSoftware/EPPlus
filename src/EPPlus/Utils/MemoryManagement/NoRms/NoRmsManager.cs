using OfficeOpenXml.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.MemoryManagement.NoRsms
{
    internal class NoRmsManager : IMemoryManager
    {
        public int DefaultBlockSize => throw new NotImplementedException();

        public int DefaultLargeBufferMultiple => throw new NotImplementedException();

        public int DefaultMaximumBufferSize => throw new NotImplementedException();

        public long DefaultMaxSmallPoolFreeBytes => throw new NotImplementedException();

        public long DefaultMaxLargePoolFreeBytes => throw new NotImplementedException();

        public MemoryStream GetStream()
        {
           return new MemoryStream();
        }

        public MemoryStream GetStream(Guid id)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag, long requiredSize)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(Guid id, string tag, long requiredSize)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(Guid id, string tag, long requiredSize, bool asContiguousBuffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag, long requiredSize, bool asContiguousBuffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(Guid id, string tag, byte[] buffer, int offset, int count)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(byte[] buffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag, byte[] buffer, int offset, int count)
        {
            return new MemoryStream();
        }

#if !NET35

        public MemoryStream GetStream(Guid id, string tag, ReadOnlySpan<byte> buffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(ReadOnlySpan<byte> buffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag, ReadOnlySpan<byte> buffer)
        {
            return new MemoryStream();
        }
#endif
    }
}
