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
		private static object _dataLock = new object();

		private static Microsoft.IO.RecyclableMemoryStreamManager MemoryManager
		{
			get
			{
				return LazyInitializer.EnsureInitialized(ref _memoryManager, ref _dataInitialized, ref _dataLock);
			}
		}

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
		/// <param name="capacity">The capacity to</param>
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
