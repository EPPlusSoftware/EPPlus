using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.Interfaces
{
    /// <summary>
    /// Memory manager interface
    /// </summary>
    public interface IMemoryManager
    {
        /// <summary>
        /// Default block size, in bytes.
        /// </summary>
        int DefaultBlockSize { get; }

        /// <summary>
        /// Default large buffer multiple, in bytes.
        /// </summary>
        int DefaultLargeBufferMultiple { get; }

        /// <summary>
        /// Default maximum buffer size, in bytes.
        /// </summary>
        int DefaultMaximumBufferSize { get; }

        MemoryStream GetStream();

        MemoryStream GetStream(Guid id);

        MemoryStream GetStream(string? tag);

        MemoryStream GetStream(string? tag, long requiredSize);

        MemoryStream GetStream(Guid id, string? tag, long requiredSize);

        MemoryStream GetStream(Guid id, string? tag, long requiredSize, bool asContiguousBuffer);

        MemoryStream GetStream(string? tag, long requiredSize, bool asContiguousBuffer);

        MemoryStream GetStream(Guid id, string? tag, byte[] buffer, int offset, int count);

        MemoryStream GetStream(byte[] buffer);

        MemoryStream GetStream(string? tag, byte[] buffer, int offset, int count);

#if !NET35
        MemoryStream GetStream(Guid id, string? tag, ReadOnlySpan<byte> buffer);

        MemoryStream GetStream(ReadOnlySpan<byte> buffer);

        MemoryStream GetStream(string? tag, ReadOnlySpan<byte> buffer);


#endif








    }
}
