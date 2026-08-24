using System;
using System.IO;

namespace OfficeOpenXml.Packaging.Ionic
{
    /// <summary>
    /// Supplies MemoryStreams to the bundled DotNetZip/Ionic compression code.
    /// Defaults to plain MemoryStreams so this assembly depends on nothing.
    /// EPPlus overrides the two providers at startup to route through its pooled
    /// RecyclableMemoryStream path. This inverts what used to be a direct call
    /// into EPPlus's internal EPPlusMemoryManager.
    /// </summary>
    public static class CompressionStreamFactory
    {
        public static Func<MemoryStream> Provider { get; set; }
            = () => new MemoryStream();

        public static Func<byte[], MemoryStream> BufferProvider { get; set; }
            = buffer => new MemoryStream(buffer);

        public static MemoryStream GetStream() => Provider();
        public static MemoryStream GetStream(byte[] buffer) => BufferProvider(buffer);
    }
}