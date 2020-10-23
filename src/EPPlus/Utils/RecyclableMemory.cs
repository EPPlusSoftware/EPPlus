#if !NET35

using Microsoft.IO;

#endif

using System.IO;

namespace OfficeOpenXml.Utils
{
	internal static class RecyclableMemory
	{
#if !NET35
		private static RecyclableMemoryStreamManager _memoryManager;

		private static RecyclableMemoryStreamManager MemoryManager
		{
			get 
			{
				if (_memoryManager is null)
				{
					_memoryManager = new RecyclableMemoryStreamManager();
				}

				return _memoryManager;
			}
		}

		public static void SetRecyclableMemoryStreamManager(RecyclableMemoryStreamManager recyclableMemoryStreamManager)
		{
			_memoryManager = recyclableMemoryStreamManager;
		}
#endif
		public static MemoryStream GetStream()
		{
#if NET35
			return new MemoryStream();
#else
			return MemoryManager.GetStream();
#endif
		}

		public static MemoryStream GetStream(byte[] array)
		{
#if NET35
			return new MemoryStream(array);
#else
			return MemoryManager.GetStream(array);
#endif
		}

		public static MemoryStream GetStream(int capacity)
		{
#if NET35
			return new MemoryStream(capacity);
#else
			return MemoryManager.GetStream(null, capacity);
#endif
		}

	}
}
