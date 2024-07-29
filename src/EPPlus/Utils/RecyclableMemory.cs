#if !NET35
using Microsoft.IO;
#endif
using System;
using System.IO;
using System.Threading;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;

namespace OfficeOpenXml.Utils
{

#if !NET35
    /// <summary>
    /// Memmory settings for RecyclableMemoryStream handling
    /// </summary>
    public class MemorySettings
	{
        /// <summary>
        /// The memory manager used, if RecyclableMemoryStream are used.
		/// <seealso cref="UseRecyclableMemory"/>
        /// </summary>
        public RecyclableMemoryStreamManager MemoryManager 
		{
			get
			{
				return RecyclableMemory.MemoryManager;
			}
			set
			{
				if (value == null)
				{
					throw new ArgumentNullException("Memory manager must not be null.");
				}
				if(RecyclableMemory.HasMemoryManager)
				{
                    throw new InvalidOperationException("A Memory Manager has already been created. To set a new memory manager, setting this property must be done before creating or opening any package.");
                }
				RecyclableMemory.SetRecyclableMemoryStreamManager(value);
			}
		}
        /// <summary>
        /// If true RecyclableMemoryStream's will be used to handle MemoryStreams. Default.
		/// If false normal MemoryStream will be used.
        /// </summary>
        public bool UseRecyclableMemory 
		{
			get
			{
				return RecyclableMemory.UseRecyclableMemory;
			}
			set
			{
				RecyclableMemory.UseRecyclableMemory = value;
            } 
		}
	}
#endif
	/// <summary>
	/// Handles the Recyclable Memory stream for supported and unsupported target frameworks.
	/// </summary>
	internal class RecyclableMemory
	{
#if !NET35
		private static RecyclableMemoryStreamManager _memoryManager;
		private static object _dataLock = new object();

		public static bool UseRecyclableMemory { get; set; } = true;
		internal static bool HasMemoryManager
		{
			get
			{
				return _memoryManager != null;
			}
		}
        internal static RecyclableMemoryStreamManager MemoryManager
		{
			get
			{
				lock(_dataLock)
                {						
					if (_memoryManager == null)
                    {
						_memoryManager = new RecyclableMemoryStreamManager();
                    }
                }
                return _memoryManager;
            }
        }

		/// <summary>
		/// Sets the RecyclableMemorytreamsManager to manage pools
		/// </summary>
		/// <param name="recyclableMemoryStreamManager">The memory manager</param>
		public static void SetRecyclableMemoryStreamManager(RecyclableMemoryStreamManager recyclableMemoryStreamManager)
		{
            _memoryManager = recyclableMemoryStreamManager;
		}
#endif
		/// <summary>
		/// Get a new memory stream.
		/// </summary>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream()
		{
#if NET35
			return new MemoryStream();
#else
			if (UseRecyclableMemory == true)
			{
				return MemoryManager.GetStream();
			}
			else
			{
                return new MemoryStream();

            }
#endif
        }

		/// <summary>
		/// Get a new memory stream initiated with a byte-array
		/// </summary>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream(byte[] array)
		{
#if NET35
			return new MemoryStream(array);
#else
            if(UseRecyclableMemory==true)
			{
                return MemoryManager.GetStream(array);
            }
            else
			{
                return new MemoryStream(array);
            }
#endif
        }
		/// <summary>
		/// Get a new memory stream initiated with a byte-array
		/// </summary>
		/// <param name="capacity">The initial size of the internal array</param>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream(int capacity)
		{
#if NET35
			return new MemoryStream(capacity);
#else
			if (UseRecyclableMemory == true)
			{
				return MemoryManager.GetStream(null, capacity);
			}
            else
            {
                return new MemoryStream(capacity);
            }
#endif
        }
    }
}
