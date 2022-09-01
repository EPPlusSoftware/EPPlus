using System;
using System.IO;
using System.Threading;

namespace OfficeOpenXml.Utils
{
	/// <summary>
	/// Handles the Recyclable Memory stream for supported and unsupported target frameworks.
	/// </summary>
	public static class RecyclableMemory
	{
#if !NET35
		private static Microsoft.IO.RecyclableMemoryStreamManager _memoryManager;
		private static bool _dataInitialized = false;
		private static bool _lazyInitializeFailed = false;
		private static object _dataLock = new object();

		private static Microsoft.IO.RecyclableMemoryStreamManager MemoryManager
		{
			get
			{
				var manager = default(Microsoft.IO.RecyclableMemoryStreamManager);
				if (_lazyInitializeFailed && _dataInitialized)
				{
					return _memoryManager;
				}
				// This has failed on dalvikvm (android), so adding a fallback handling of Exceptions /MA 2022-08-31
				try
                {
					manager = LazyInitializer.EnsureInitialized(ref _memoryManager, ref _dataInitialized, ref _dataLock);
				}
				catch(Exception)
                {
					lock(_dataLock)
                    {
						_lazyInitializeFailed = true;
						if (_memoryManager == null)
                        {
							_memoryManager = new Microsoft.IO.RecyclableMemoryStreamManager();
							_dataInitialized = true;
                        }
                    }
					manager = _memoryManager;
				}
				return manager;
			}
		}

		/// <summary>
		/// Sets the RecyclableMemorytreamsManager to manage pools
		/// </summary>
		/// <param name="recyclableMemoryStreamManager">The memory manager</param>
		public static void SetRecyclableMemoryStreamManager(Microsoft.IO.RecyclableMemoryStreamManager recyclableMemoryStreamManager)
		{
			_dataInitialized = recyclableMemoryStreamManager is object;
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
			return MemoryManager.GetStream();
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
			return MemoryManager.GetStream(array);
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
			return MemoryManager.GetStream(null, capacity);
#endif
		}
	}
}
