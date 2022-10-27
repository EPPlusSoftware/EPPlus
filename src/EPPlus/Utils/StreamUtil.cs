/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/27/2022         EPPlus Software AB       Initial release EPPlus 6.1
 *************************************************************************************************/
using System;
using System.IO;
using System.Threading;
#if !NET35
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml.Utils
{
    internal class StreamUtil
    {
        static object _lock = new object();
        /// <summary>
        /// Copies the input stream to the output stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="outputStream">The output stream.</param>
        internal static void CopyStream(Stream inputStream, ref Stream outputStream)
        {
            if (!inputStream.CanRead)
            {
                throw (new Exception("Cannot read from the input stream"));
            }
            if (!outputStream.CanWrite)
            {
                throw (new Exception("Cannot write to the output stream"));
            }
            if (inputStream.CanSeek)
            {
                inputStream.Seek(0, SeekOrigin.Begin);
            }

            const int bufferLength = 8096;
            var buffer = new Byte[bufferLength];
            lock (_lock)
            {
                int bytesRead = inputStream.Read(buffer, 0, bufferLength);
                // write the required bytes
                while (bytesRead > 0)
                {
                    outputStream.Write(buffer, 0, bytesRead);
                    bytesRead = inputStream.Read(buffer, 0, bufferLength);
                }
                outputStream.Flush();
            }
        }
        #if !NET35
        /// <summary>
        /// Copies the input stream to the output stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="outputStream">The output stream.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        internal static async Task CopyStreamAsync(Stream inputStream, Stream outputStream, CancellationToken cancellationToken)
        {
            if (!inputStream.CanRead)
            {
                throw new Exception("Cannot read from the input stream");
            }
            if (!outputStream.CanWrite)
            {
                throw new Exception("Cannot write to the output stream");
            }
            if (inputStream.CanSeek)
            {
                inputStream.Seek(0, SeekOrigin.Begin);
            }

            const int bufferLength = 8096;
            var buffer = new byte[bufferLength];
            var bytesRead = await inputStream.ReadAsync(buffer, 0, bufferLength, cancellationToken).ConfigureAwait(false);
            // write the required bytes
            while (bytesRead > 0)
            {
                await outputStream.WriteAsync(buffer, 0, bytesRead, cancellationToken).ConfigureAwait(false);
                bytesRead = await inputStream.ReadAsync(buffer, 0, bufferLength, cancellationToken).ConfigureAwait(false);
            }
            await outputStream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        #endif
    }
}
