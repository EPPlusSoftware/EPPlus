/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/23/2025         EPPlus Software AB           ArrayPoolHelper implementation
 *************************************************************************************************/
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
using System.Buffers;
#endif
using System;

namespace EPPlus.Fonts.OpenType.Utilities
{
    /// <summary>
    /// Helper class for array pooling with fallback to regular allocation on older frameworks.
    /// Provides a consistent API across all .NET versions.
    /// </summary>
    internal static class ArrayPoolHelper<T>
    {
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
        private static readonly ArrayPool<T> Pool = ArrayPool<T>.Shared;
#endif
        // Cache empty array for .NET 3.5 compatibility (Array.Empty<T>() was added in .NET 4.6)
        private static readonly T[] EmptyArray = new T[0];

        /// <summary>
        /// Rents an array from the pool (or allocates a new one in older targets).
        /// The returned array may be larger than the requested minimum length.
        /// </summary>
        /// <param name="minimumLength">Minimum required array length</param>
        /// <returns>An array with at least minimumLength elements</returns>
        public static T[] Rent(int minimumLength)
        {
            if (minimumLength < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(minimumLength), "Minimum length must be non-negative");
            }

            if (minimumLength == 0)
            {
                return EmptyArray;
            }

#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
            return Pool.Rent(minimumLength);
#else
            return new T[minimumLength];
#endif
        }

        /// <summary>
        /// Returns the array to the pool (does nothing in older targets).
        /// </summary>
        /// <param name="array">Array to return</param>
        /// <param name="clearArray">If true, clears the array before returning to pool</param>
        public static void Return(T[] array, bool clearArray = false)
        {
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
            if (array != null && array.Length > 0)
            {
                Pool.Return(array, clearArray);
            }
#endif
            // In older targets: let GC handle the array
        }

        /// <summary>
        /// Safely returns the array to the pool and sets the reference to null.
        /// This prevents accidental reuse of returned arrays.
        /// </summary>
        /// <param name="array">Reference to array to return</param>
        /// <param name="clearArray">If true, clears the array before returning to pool</param>
        public static void SafeReturn(ref T[] array, bool clearArray = false)
        {
            if (array != null && array.Length > 0)
            {
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
                Pool.Return(array, clearArray);
#endif
                array = null;
            }
        }

        /// <summary>
        /// Ensures an array has at least the specified capacity.
        /// If the current array is too small, returns it to the pool and rents a larger one.
        /// If the current array is sufficient, returns it unchanged.
        /// </summary>
        /// <param name="array">Current array (may be null)</param>
        /// <param name="currentCapacity">Tracked capacity of current array</param>
        /// <param name="minimumLength">Required minimum length</param>
        /// <param name="clearArray">If true and a new array is rented, clears it</param>
        /// <returns>Array with at least minimumLength capacity</returns>
        public static T[] EnsureCapacity(ref T[] array, ref int currentCapacity, int minimumLength, bool clearArray = false)
        {
            if (minimumLength < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(minimumLength), "Minimum length must be non-negative");
            }

            if (minimumLength == 0)
            {
                return EmptyArray;
            }

            // If we already have a sufficient array, return it
            if (array != null && currentCapacity >= minimumLength)
            {
                return array;
            }

            // Return old array to pool
            if (array != null && array.Length > 0)
            {
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
                Pool.Return(array, clearArray: false);
#endif
            }

            // Rent new array
#if NETSTANDARD2_1 || NETCOREAPP2_1_OR_GREATER || NET5_0_OR_GREATER
            array = Pool.Rent(minimumLength);
            currentCapacity = array.Length;

            if (clearArray)
            {
                Array.Clear(array, 0, array.Length);
            }
#else
            array = new T[minimumLength];
            currentCapacity = minimumLength;
#endif

            return array;
        }

        /// <summary>
        /// Rents an array and copies data from source array.
        /// Useful for resizing operations.
        /// </summary>
        /// <param name="source">Source array to copy from</param>
        /// <param name="sourceLength">Number of elements to copy from source</param>
        /// <param name="minimumLength">Minimum length of new array (must be >= sourceLength)</param>
        /// <returns>New array with copied data</returns>
        public static T[] RentAndCopy(T[] source, int sourceLength, int minimumLength)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            if (sourceLength < 0 || sourceLength > source.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(sourceLength));
            }

            if (minimumLength < sourceLength)
            {
                throw new ArgumentException("Minimum length must be at least sourceLength", nameof(minimumLength));
            }

            var newArray = Rent(minimumLength);

            if (sourceLength > 0)
            {
                Array.Copy(source, 0, newArray, 0, sourceLength);
            }

            return newArray;
        }

        /// <summary>
        /// Creates a scope that automatically returns the array when disposed.
        /// Usage: using (var scope = ArrayPoolHelper{T}.RentScoped(100)) { ... }
        /// </summary>
        /// <param name="minimumLength">Minimum required array length</param>
        /// <param name="clearOnReturn">If true, clears the array before returning to pool</param>
        /// <returns>A disposable scope containing the rented array</returns>
        public static RentedArrayScope RentScoped(int minimumLength, bool clearOnReturn = false)
        {
            return new RentedArrayScope(Rent(minimumLength), clearOnReturn);
        }

        /// <summary>
        /// Disposable wrapper for rented arrays that ensures they are returned to the pool.
        /// </summary>
        public struct RentedArrayScope : IDisposable
        {
            private T[] _array;
            private readonly bool _clearOnReturn;
            private bool _disposed;

            internal RentedArrayScope(T[] array, bool clearOnReturn)
            {
                _array = array;
                _clearOnReturn = clearOnReturn;
                _disposed = false;
            }

            /// <summary>
            /// Gets the rented array.
            /// </summary>
            public T[] Array
            {
                get
                {
                    if (_disposed)
                    {
                        throw new ObjectDisposedException(nameof(RentedArrayScope));
                    }
                    return _array;
                }
            }

            /// <summary>
            /// Gets the length of the rented array.
            /// </summary>
            public int Length => _array?.Length ?? 0;

            /// <summary>
            /// Returns the array to the pool.
            /// </summary>
            public void Dispose()
            {
                if (!_disposed && _array != null)
                {
                    ArrayPoolHelper<T>.Return(_array, _clearOnReturn);
                    _array = null;
                    _disposed = true;
                }
            }
        }
    }
}