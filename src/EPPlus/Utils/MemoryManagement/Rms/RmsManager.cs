#if !NET35
using Microsoft.IO;
#endif
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.MemoryManagement.Rms
{
    internal class RmsManager : IMemoryManager
    {
#if !NET35
        public RmsManager()
        {

            var defaultOptions = new RecyclableMemoryStreamManager.Options
            {
                MaximumLargePoolFreeBytes = 32 * 1024 * 1024,
                MaximumSmallPoolFreeBytes = 512 * 1024,
            };
            _memoryStreamManager = new RecyclableMemoryStreamManager(defaultOptions);
        }

        private RecyclableMemoryStreamManager _memoryStreamManager;

        public int DefaultBlockSize => RecyclableMemoryStreamManager.DefaultBlockSize;

        public int DefaultLargeBufferMultiple => RecyclableMemoryStreamManager.DefaultLargeBufferMultiple;

        public int DefaultMaximumBufferSize => RecyclableMemoryStreamManager.DefaultMaximumBufferSize;

        public MemoryStream GetStream()
        {
            return _memoryStreamManager.GetStream();
        }

        public MemoryStream GetStream(Guid id)
        {
            return _memoryStreamManager.GetStream(id);
        }

        public MemoryStream GetStream(string tag)
        {
            return _memoryStreamManager.GetStream(tag);
        }

        public MemoryStream GetStream(string tag, long requiredSize)
        {
            return _memoryStreamManager.GetStream(tag, requiredSize);
        }

        public MemoryStream GetStream(Guid id, string tag, long requiredSize)
        {
            return _memoryStreamManager.GetStream(id, tag, requiredSize);
        }

        public MemoryStream GetStream(Guid id, string tag, long requiredSize, bool asContiguousBuffer)
        {
            return _memoryStreamManager.GetStream(id, tag, requiredSize, asContiguousBuffer);
        }

        public MemoryStream GetStream(string tag, long requiredSize, bool asContiguousBuffer)
        {
            return _memoryStreamManager.GetStream(tag, requiredSize, asContiguousBuffer);
        }

        public MemoryStream GetStream(Guid id, string tag, byte[] buffer, int offset, int count)
        {
            return _memoryStreamManager.GetStream(id, tag, buffer, offset, count);
        }

        public MemoryStream GetStream(byte[] buffer)
        {
            return _memoryStreamManager.GetStream(buffer);
        }

        public MemoryStream GetStream(string tag, byte[] buffer, int offset, int count)
        {
            return _memoryStreamManager.GetStream(tag, buffer, offset, count);
        }

        public MemoryStream GetStream(Guid id, string tag, ReadOnlySpan<byte> buffer)
        {
            return _memoryStreamManager.GetStream(id, tag, buffer);
        }

        public MemoryStream GetStream(ReadOnlySpan<byte> buffer)
        {
            return _memoryStreamManager.GetStream(buffer);
        }

        public MemoryStream GetStream(string tag, ReadOnlySpan<byte> buffer)
        {
            return _memoryStreamManager.GetStream(tag, buffer);
        }
#else
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

        public MemoryStream GetStream(byte[] buffer)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(Guid id, string? tag, byte[] buffer, int offset, int count)
        {
             return new MemoryStream();
        }

         public MemoryStream GetStream(string tag, byte[] buffer, int offset, int count)
        {
            return new MemoryStream();
        }

        public MemoryStream GetStream(string tag, long requiredSize, bool asContiguousBuffer)
        {
            return new MemoryStream();
        }
#endif


    }
}
