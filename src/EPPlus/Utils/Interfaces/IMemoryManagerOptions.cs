using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.Interfaces
{
    /// <summary>
    /// MemoryManagerOptions interface
    /// </summary>
    public class IMemoryManagerOptions
    {
        /// <summary>
        /// Gets or sets the size of the pooled blocks. This must be greater than 0.
        /// </summary>
        /// <remarks>The default size 131,072 (128KB)</remarks>
        int BlockSize { get; set; }

        /// <summary>
        /// Each large buffer will be a multiple exponential of this value
        /// </summary>
        /// <remarks>The default value is 1,048,576 (1MB)</remarks>
        int LargeBufferMultiple { get; set; }

        /// <summary>
        /// Buffer beyond this length are not pooled.
        /// </summary>
        /// <remarks>The default value is 134,217,728 (128MB)</remarks>
        int MaximumBufferSize { get; set; }

        /// <summary>
        /// Maximum number of bytes to keep available in the small pool.
        /// </summary>
        /// <remarks>
        /// <para>Trying to return buffers to the pool beyond this limit will result in them being garbage collected.</para>
        /// <para>The default value is 0, but all users should set a reasonable value depending on your application's memory requirements.</para>
        /// </remarks>
        long MaximumSmallPoolFreeBytes { get; set; }

        /// <summary>
        /// Maximum number of bytes to keep available in the large pools.
        /// </summary>
        /// <remarks>
        /// <para>Trying to return buffers to the pool beyond this limit will result in them being garbage collected.</para>
        /// <para>The default value is 0, but all users should set a reasonable value depending on your application's memory requirements.</para>
        /// </remarks>
        long MaximumLargePoolFreeBytes { get; set; }

        /// <summary>
        /// Whether to use the exponential allocation strategy (see documentation).
        /// </summary>
        /// <remarks>The default value is false.</remarks>
        public bool UseExponentialLargeBuffer { get; set; }

        /// <summary>
        /// Maximum stream capacity in bytes. Attempts to set a larger capacity will
        /// result in an exception.
        /// </summary>
        /// <remarks>The default value of 0 indicates no limit.</remarks>
        public long MaximumStreamCapacity { get; set; }

        /// <summary>
        /// Whether to save call stacks for stream allocations. This can help in debugging.
        /// It should NEVER be turned on generally in production.
        /// </summary>
        bool GenerateCallStacks { get; set; }

        /// <summary>
        /// Whether dirty buffers can be immediately returned to the buffer pool.
        /// </summary>
        /// <remarks>
        /// <para>
        /// When <see cref="RecyclableMemoryStream.GetBuffer"/> is called on a stream and creates a single large buffer, if this setting is enabled, the other blocks will be returned
        /// to the buffer pool immediately.
        /// </para>
        /// <para>
        /// Note when enabling this setting that the user is responsible for ensuring that any buffer previously
        /// retrieved from a stream which is subsequently modified is not used after modification (as it may no longer
        /// be valid).
        /// </para>
        /// </remarks>
        bool AggressiveBufferReturn { get; set; }

        /// <summary>
        /// Causes an exception to be thrown if <see cref="RecyclableMemoryStream.ToArray"/> is ever called.
        /// </summary>
        /// <remarks>Calling <see cref="RecyclableMemoryStream.ToArray"/> defeats the purpose of a pooled buffer. Use this property to discover code that is calling <see cref="RecyclableMemoryStream.ToArray"/>. If this is
        /// set and <see cref="RecyclableMemoryStream.ToArray"/> is called, a <c>NotSupportedException</c> will be thrown.</remarks>
        bool ThrowExceptionOnToArray { get; set; }

        /// <summary>
        /// Zero out buffers on allocation and before returning them to the pool.
        /// </summary>
        /// <remarks>Setting this to true causes a performance hit and should only be set if one wants to avoid accidental data leaks.</remarks>
        bool ZeroOutBuffer { get; set; }
    }
}
