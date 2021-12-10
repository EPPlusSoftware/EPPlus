/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.IO;
#if !NET35 && !NET40
using System.Threading;
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml
{
    public sealed partial class ExcelPackage
    {
#if !NET35 && !NET40
        #region Load

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="fileInfo">The input file.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(FileInfo fileInfo, CancellationToken cancellationToken = default)
        {
            var stream = fileInfo.OpenRead();
            await LoadAsync(stream, RecyclableMemory.GetStream(), null, cancellationToken).ConfigureAwait(false);
        }
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, CancellationToken cancellationToken = default)
        {
            await LoadAsync(new FileInfo(filePath), cancellationToken);
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="fileInfo">The input file.</param>
        /// <param name="Password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(FileInfo fileInfo, string Password, CancellationToken cancellationToken = default)
        {
            var stream = fileInfo.OpenRead();
            await LoadAsync(stream, RecyclableMemory.GetStream(), Password, cancellationToken).ConfigureAwait(false);
        }
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, string password, CancellationToken cancellationToken = default)
        {
            await LoadAsync(new FileInfo(filePath), password, cancellationToken);
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="fileInfo">The input file.</param>
        /// <param name="output">The out stream. Sets the Stream property</param>
        /// <param name="Password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(FileInfo fileInfo, Stream output, string Password, CancellationToken cancellationToken = default)
        {
            var stream = fileInfo.OpenRead();
            await LoadAsync(stream, output, Password, cancellationToken).ConfigureAwait(false);
        }
        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="output">The out stream. Sets the Stream property</param>
        /// <param name="password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, Stream output, string password, CancellationToken cancellationToken = default)
        {
            await LoadAsync(new FileInfo(filePath), output, password, cancellationToken);
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(Stream input, CancellationToken cancellationToken = default)
        {
            await LoadAsync(input, RecyclableMemory.GetStream(), null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="Password">The password to decrypt the document</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(Stream input, string Password, CancellationToken cancellationToken = default)
        {
            await LoadAsync(input, RecyclableMemory.GetStream(), Password, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>    
        /// <param name="output"></param>
        /// <param name="Password"></param>
        /// <param name="cancellationToken"></param>
        private async Task LoadAsync(Stream input, Stream output, string Password, CancellationToken cancellationToken)
        {
            ReleaseResources();
            if (input.CanSeek && input.Length == 0) // Template is blank, Construct new
            {
                _stream = output;
                await ConstructNewFileAsync(Password, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                Stream ms;
                _stream = output;
                if (Password != null)
                {
                    using (var encrStream = RecyclableMemory.GetStream())
                    {
                        await CopyStreamAsync(input, encrStream, cancellationToken).ConfigureAwait(false);
                        var eph = new EncryptedPackageHandler();
                        Encryption.Password = Password;
                        ms = eph.DecryptPackage(encrStream, Encryption);
                    }
                }
                else
                {
                    ms = RecyclableMemory.GetStream();
                    await CopyStreamAsync(input, ms, cancellationToken).ConfigureAwait(false);
                }

				try
				{
					_zipPackage = new Packaging.ZipPackage(ms);
				}
				catch (Exception ex)
				{
					if (Password == null && await CompoundDocumentFile.IsCompoundDocumentAsync((MemoryStream)_stream, cancellationToken).ConfigureAwait(false))
					{
						throw new Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
					}

					throw;
				}
                finally
                {
                    ms.Dispose();
				}
            }
            //Clear the workbook so that it gets reinitialized next time
            this._workbook = null;
        }
        #endregion

        #region SaveAsync

        /// <summary>
        /// Saves all the components back into the package.
        /// This method recursively calls the Save method on all sub-components.
        /// The package is closed after it has ben saved
        /// d to encrypt the workbook with. 
        /// </summary>
        /// <returns></returns>
        public async Task SaveAsync(CancellationToken cancellationToken = default)
        {
            CheckNotDisposed();
            try
            {
                if (_stream is MemoryStream && _stream.Length > 0)
                {
                    //Close any open memorystream and "renew" then. This can occure if the package is saved twice. 
                    //The stream is left open on save to enable the user to read the stream-property.
                    //Non-memorystream streams will leave the closing to the user before saving a second time.
                    CloseStream();
                }

                //Invoke before save delegates
                foreach (var action in BeforeSave)
                {
                    action.Invoke();
                }

                Workbook.Save();
                if (File == null)
                {
                    if (Encryption.IsEncrypted)
                    {
                        using (var ms = RecyclableMemory.GetStream())
                        {
                            _zipPackage.Save(ms);
                            var file = ms.ToArray();
                            var eph = new EncryptedPackageHandler();
                            using (var msEnc = eph.EncryptPackage(file, Encryption))
                            {
                                await CopyStreamAsync(msEnc, _stream, cancellationToken).ConfigureAwait(false);
                            }
                        }
                    }
                    else
                    {
                        _zipPackage.Save(_stream);
                    }
                    await _stream.FlushAsync(cancellationToken);
                    _zipPackage.Close();
                }
                else
                {
                    if (System.IO.File.Exists(File.FullName))
                    {
                        try
                        {
                            System.IO.File.Delete(File.FullName);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Error overwriting file {File.FullName}", ex);
                        }
                    }

                    _zipPackage.Save(_stream);
                    _zipPackage.Close();
                    if (Stream is MemoryStream stream)
                    {
#if NETSTANDARD2_1
                        await using (var fi = new FileStream(File.FullName, FileMode.Create))
#else
                        using (var fi = new FileStream(File.FullName, FileMode.Create))
#endif
                        {
                            //EncryptPackage
                            if (Encryption.IsEncrypted)
                            {
                                var file = stream.ToArray();
                                var eph = new EncryptedPackageHandler();
                                using (var ms = eph.EncryptPackage(file, Encryption))
                                {
                                    await fi.WriteAsync(ms.ToArray(), 0, (int)ms.Length, cancellationToken).ConfigureAwait(false);
                                }
                            }
                            else
                            {
                                await fi.WriteAsync(stream.ToArray(), 0, (int)Stream.Length, cancellationToken).ConfigureAwait(false);
                            }
                        }
                    }
                    else
                    {
#if NETSTANDARD2_1
                        await using (var fs = new FileStream(File.FullName, FileMode.Create))
#else
                        using (var fs = new FileStream(File.FullName, FileMode.Create))
#endif
                        {
                            var b = await GetAsByteArrayAsync(false, cancellationToken).ConfigureAwait(false);
                            await fs.WriteAsync(b, 0, b.Length, cancellationToken).ConfigureAwait(false);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                if (File == null)
                {
                    throw;
                }

                throw new InvalidOperationException($"Error saving file {File.FullName}", ex);
            }
        }

        /// <summary>
        /// Saves all the components back into the package.
        /// This method recursively calls the Save method on all sub-components.
        /// The package is closed after it has ben saved
        /// Supply a password to encrypt the workbook package. 
        /// </summary>
        /// <param name="password">This parameter overrides the Workbook.Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsync(string password, CancellationToken cancellationToken = default)
        {
            Encryption.Password = password;
            await SaveAsync(cancellationToken).ConfigureAwait(false);
        }

        #endregion

        #region SaveAsAsync

        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved        
        /// </summary>
        /// <param name="file">The file location</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(FileInfo file, CancellationToken cancellationToken = default)
        {
            File = file;
            await SaveAsync(cancellationToken).ConfigureAwait(false); 
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved        
        /// </summary>
        /// <param name="filePath">The file location</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(string filePath, CancellationToken cancellationToken = default)
        {
            await SaveAsAsync(new FileInfo(filePath), cancellationToken);
        }

        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="file">The file</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(FileInfo file, string password, CancellationToken cancellationToken = default)
        {
            File = file;
            Encryption.Password = password;
            await SaveAsync(cancellationToken).ConfigureAwait(false);
        }
        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="filePath">The file</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(string filePath, string password, CancellationToken cancellationToken = default)
        {
            await SaveAsAsync(new FileInfo(filePath), password, cancellationToken);
        }

        /// <summary>
        /// Copies the Package to the Outstream
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="OutputStream">The stream to copy the package to</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(Stream OutputStream, CancellationToken cancellationToken = default)
        {
            File = null;
            await SaveAsync(cancellationToken).ConfigureAwait(false); 

            if (OutputStream != _stream)
            {
                await CopyStreamAsync(_stream, OutputStream, cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Copies the Package to the Outstream
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="OutputStream">The stream to copy the package to</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(Stream OutputStream, string password, CancellationToken cancellationToken = default)
        {
            Encryption.Password = password;
            await SaveAsAsync(OutputStream, cancellationToken).ConfigureAwait(false); 
        }

        #endregion

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

        internal async Task<byte[]> GetAsByteArrayAsync(bool save, CancellationToken cancellationToken)
        {
            CheckNotDisposed();
            if (save)
            {
                Workbook.Save();
                _zipPackage.Close();
                if (_stream is MemoryStream && _stream.Length > 0)
                {
                    _stream.Close();
#if Standard21
                    await _stream.DisposeAsync();
#else
                    _stream.Dispose();
#endif
                    _stream = RecyclableMemory.GetStream();
                }
                _zipPackage.Save(_stream);
            }
            var byRet = new byte[Stream.Length];
            var pos = Stream.Position;
            Stream.Seek(0, SeekOrigin.Begin);
            await Stream.ReadAsync(byRet, 0, (int)Stream.Length, cancellationToken).ConfigureAwait(false);

            //Encrypt Workbook?
            if (Encryption.IsEncrypted)
            {
                var eph = new EncryptedPackageHandler();
                using (var ms = eph.EncryptPackage(byRet, Encryption))
                {
                    byRet = ms.ToArray();
                }
            }

            Stream.Seek(pos, SeekOrigin.Begin);
            Stream.Close();
            return byRet;
        }

        /// <summary>
        /// Saves and returns the Excel files as a bytearray.
        /// Note that the package is closed upon save
        /// </summary>
        /// <example>      
        /// Example how to return a document from a Webserver...
        /// <code> 
        ///  ExcelPackage package=new ExcelPackage();
        ///  /**** ... Create the document ****/
        ///  Byte[] bin = package.GetAsByteArray();
        ///  Response.ContentType = "Application/vnd.ms-Excel";
        ///  Response.AddHeader("content-disposition", "attachment;  filename=TheFile.xlsx");
        ///  Response.BinaryWrite(bin);
        /// </code>
        /// </example>
        /// <param name="cancellationToken">The cancellation token</param>
        /// <returns></returns>
        public async Task<byte[]> GetAsByteArrayAsync(CancellationToken cancellationToken = default)
        {
            return await GetAsByteArrayAsync(true, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Saves and returns the Excel files as a bytearray
        /// Note that the package is closed upon save
        /// </summary>
        /// <example>      
        /// Example how to return a document from a Webserver...
        /// <code> 
        ///  ExcelPackage package=new ExcelPackage();
        ///  /**** ... Create the document ****/
        ///  Byte[] bin = package.GetAsByteArray();
        ///  Response.ContentType = "Application/vnd.ms-Excel";
        ///  Response.AddHeader("content-disposition", "attachment;  filename=TheFile.xlsx");
        ///  Response.BinaryWrite(bin);
        /// </code>
        /// </example>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        /// <returns></returns>
        public async Task<byte[]> GetAsByteArrayAsync(string password, CancellationToken cancellationToken = default)
        {
            if (password != null)
            {
                Encryption.Password = password;
            }
            return await GetAsByteArrayAsync(true, cancellationToken).ConfigureAwait(false);
        }

        private async Task ConstructNewFileAsync(string password, CancellationToken cancellationToken)
        {
            var ms = RecyclableMemory.GetStream();
            if (_stream == null) _stream = RecyclableMemory.GetStream();
            File?.Refresh();
            if (File != null && File.Exists)
            {
                if (password != null)
                {
                    var encrHandler = new EncryptedPackageHandler();
                    Encryption.IsEncrypted = true;
                    Encryption.Password = password;
                    ms.Dispose();
                    ms = encrHandler.DecryptPackage(File, Encryption);
                }
                else
                {
                    await WriteFileToStreamAsync(File.FullName, ms, cancellationToken).ConfigureAwait(false);
                }
                try
                {
                    _zipPackage = new Packaging.ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    if (password == null && await CompoundDocumentFile.IsCompoundDocumentAsync(File, cancellationToken).ConfigureAwait(false))
                    {
                        throw new Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password", ex);
                    }

                    throw;
                }
                finally
                {
                    ms.Dispose();
				}
            }
            else
            {
                _zipPackage = new Packaging.ZipPackage(ms);
                ms.Dispose();
                CreateBlankWb();
            }
        }

        private static async Task WriteFileToStreamAsync(string path, Stream stream,
            CancellationToken cancellationToken)
        {
#if NETSTANDARD2_1
            await using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
#else
            using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
#endif
            {
                var buffer = new byte[4096];
                int read;
                while ((read = await fileStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
                {
                    await stream.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                }
            }
        }

#endif
                }
            }
