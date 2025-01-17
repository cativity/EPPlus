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
            using FileStream? stream = fileInfo.OpenRead();
            await this.LoadAsync(stream, RecyclableMemory.GetStream(), null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, CancellationToken cancellationToken = default) => await this.LoadAsync(new FileInfo(filePath), cancellationToken);

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="fileInfo">The input file.</param>
        /// <param name="Password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(FileInfo fileInfo, string Password, CancellationToken cancellationToken = default)
        {
            using FileStream? stream = fileInfo.OpenRead();
            await this.LoadAsync(stream, RecyclableMemory.GetStream(), Password, cancellationToken).ConfigureAwait(false);
            stream.Close();
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, string password, CancellationToken cancellationToken = default) => await this.LoadAsync(new FileInfo(filePath), password, cancellationToken);

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="fileInfo">The input file.</param>
        /// <param name="output">The out stream. Sets the Stream property</param>
        /// <param name="Password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(FileInfo fileInfo, Stream output, string Password, CancellationToken cancellationToken = default)
        {
            using FileStream? stream = fileInfo.OpenRead();
            await this.LoadAsync(stream, output, Password, cancellationToken).ConfigureAwait(false);
            stream.Close();
        }

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="filePath">The input file.</param>
        /// <param name="output">The out stream. Sets the Stream property</param>
        /// <param name="password">The password</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(string filePath, Stream output, string password, CancellationToken cancellationToken = default) => await this.LoadAsync(new FileInfo(filePath), output, password, cancellationToken);

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(Stream input, CancellationToken cancellationToken = default) => await this.LoadAsync(input, RecyclableMemory.GetStream(), null, cancellationToken).ConfigureAwait(false);

        /// <summary>
        /// Loads the specified package data from a stream.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <param name="Password">The password to decrypt the document</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task LoadAsync(Stream input, string Password, CancellationToken cancellationToken = default) => await this.LoadAsync(input, RecyclableMemory.GetStream(), Password, cancellationToken).ConfigureAwait(false);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>    
        /// <param name="output"></param>
        /// <param name="Password"></param>
        /// <param name="cancellationToken"></param>
        private async Task LoadAsync(Stream input, Stream output, string Password, CancellationToken cancellationToken)
        {
            this.ReleaseResources();

            if (input.CanSeek && input.Length == 0) // Template is blank, Construct new
            {
                this._stream = output;
                await this.ConstructNewFileAsync(Password, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                Stream ms;
                this._stream = output;

                if (Password != null)
                {
                    using MemoryStream? encrStream = RecyclableMemory.GetStream();
                    await StreamUtil.CopyStreamAsync(input, encrStream, cancellationToken).ConfigureAwait(false);
                    EncryptedPackageHandler? eph = new EncryptedPackageHandler();
                    this.Encryption.Password = Password;
                    ms = eph.DecryptPackage(encrStream, this.Encryption);
                }
                else
                {
                    ms = RecyclableMemory.GetStream();
                    await StreamUtil.CopyStreamAsync(input, ms, cancellationToken).ConfigureAwait(false);
                }

                try
                {
                    this._zipPackage = new Packaging.ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    if (Password == null
                        && await CompoundDocumentFile.IsCompoundDocumentAsync((MemoryStream)this._stream, cancellationToken).ConfigureAwait(false))
                    {
                        throw new
                            Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password",
                                      ex);
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
            this.CheckNotDisposed();

            try
            {
                if (this._stream is MemoryStream && this._stream.Length > 0)
                {
                    //Close any open memorystream and "renew" then. This can occure if the package is saved twice. 
                    //The stream is left open on save to enable the user to read the stream-property.
                    //Non-memorystream streams will leave the closing to the user before saving a second time.
                    this.CloseStream();
                }

                //Invoke before save delegates
                foreach (Action? action in this.BeforeSave)
                {
                    action.Invoke();
                }

                this.Workbook.Save();

                if (this.File == null)
                {
                    if (this.Encryption.IsEncrypted)
                    {
                        using MemoryStream? ms = RecyclableMemory.GetStream();
                        this._zipPackage.Save(ms);
                        byte[]? file = ms.ToArray();
                        EncryptedPackageHandler? eph = new EncryptedPackageHandler();
                        using MemoryStream? msEnc = eph.EncryptPackage(file, this.Encryption);
                        await StreamUtil.CopyStreamAsync(msEnc, this._stream, cancellationToken).ConfigureAwait(false);
                    }
                    else
                    {
                        this._zipPackage.Save(this._stream);
                    }

                    await this._stream.FlushAsync(cancellationToken);
                    Packaging.ZipPackage.Close();
                }
                else
                {
                    if (System.IO.File.Exists(this.File.FullName))
                    {
                        try
                        {
                            System.IO.File.Delete(this.File.FullName);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Error overwriting file {this.File.FullName}", ex);
                        }
                    }

                    this._zipPackage.Save(this._stream);
                    Packaging.ZipPackage.Close();

                    if (this.Stream is MemoryStream stream)
                    {
#if NETSTANDARD2_1
                        await using (var fi = new FileStream(File.FullName, FileMode.Create))
#else
                        using FileStream? fi = new FileStream(this.File.FullName, FileMode.Create);
#endif

                        //EncryptPackage
                        if (this.Encryption.IsEncrypted)
                        {
                            byte[]? file = stream.ToArray();
                            EncryptedPackageHandler? eph = new EncryptedPackageHandler();
                            using MemoryStream? ms = eph.EncryptPackage(file, this.Encryption);
                            await fi.WriteAsync(ms.ToArray(), 0, (int)ms.Length, cancellationToken).ConfigureAwait(false);
                        }
                        else
                        {
                            await fi.WriteAsync(stream.ToArray(), 0, (int)this.Stream.Length, cancellationToken).ConfigureAwait(false);
                        }
                    }
                    else
                    {
#if NETSTANDARD2_1
                        await using (var fs = new FileStream(File.FullName, FileMode.Create))
#else
                        using FileStream? fs = new FileStream(this.File.FullName, FileMode.Create);
#endif
                        byte[]? b = await this.GetAsByteArrayAsync(false, cancellationToken).ConfigureAwait(false);
                        await fs.WriteAsync(b, 0, b.Length, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                if (this.File == null)
                {
                    throw;
                }

                throw new InvalidOperationException($"Error saving file {this.File.FullName}", ex);
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
            this.Encryption.Password = password;
            await this.SaveAsync(cancellationToken).ConfigureAwait(false);
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
            this.File = file;
            await this.SaveAsync(cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved        
        /// </summary>
        /// <param name="filePath">The file location</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(string filePath, CancellationToken cancellationToken = default) => await this.SaveAsAsync(new FileInfo(filePath), cancellationToken);

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
            this.File = file;
            this.Encryption.Password = password;
            await this.SaveAsync(cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Saves the workbook to a new file
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="filePath">The file</param>
        /// <param name="password">The password to encrypt the workbook with. 
        /// This parameter overrides the Encryption.Password.</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(string filePath, string password, CancellationToken cancellationToken = default) => await this.SaveAsAsync(new FileInfo(filePath), password, cancellationToken);

        /// <summary>
        /// Copies the Package to the Outstream
        /// The package is closed after it has been saved
        /// </summary>
        /// <param name="OutputStream">The stream to copy the package to</param>
        /// <param name="cancellationToken">The cancellation token</param>
        public async Task SaveAsAsync(Stream OutputStream, CancellationToken cancellationToken = default)
        {
            this.File = null;
            await this.SaveAsync(cancellationToken).ConfigureAwait(false);

            if (OutputStream != this._stream)
            {
                await StreamUtil.CopyStreamAsync(this._stream, OutputStream, cancellationToken).ConfigureAwait(false);
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
            this.Encryption.Password = password;
            await this.SaveAsAsync(OutputStream, cancellationToken).ConfigureAwait(false);
        }

        #endregion

        internal async Task<byte[]> GetAsByteArrayAsync(bool save, CancellationToken cancellationToken)
        {
            this.CheckNotDisposed();

            if (save)
            {
                this.Workbook.Save();
                Packaging.ZipPackage.Close();

                if (this._stream is MemoryStream && this._stream.Length > 0)
                {
                    this._stream.Close();
#if Standard21
                    await _stream.DisposeAsync();
#else
                    this._stream.Dispose();
#endif
                    this._stream = RecyclableMemory.GetStream();
                }

                this._zipPackage.Save(this._stream);
            }

            byte[]? byRet = new byte[this.Stream.Length];
            long pos = this.Stream.Position;
            this.Stream.Seek(0, SeekOrigin.Begin);
            await this.Stream.ReadAsync(byRet, 0, (int)this.Stream.Length, cancellationToken).ConfigureAwait(false);

            //Encrypt Workbook?
            if (this.Encryption.IsEncrypted)
            {
                EncryptedPackageHandler? eph = new EncryptedPackageHandler();
                using MemoryStream? ms = eph.EncryptPackage(byRet, this.Encryption);
                byRet = ms.ToArray();
            }

            this.Stream.Seek(pos, SeekOrigin.Begin);
            this.Stream.Close();

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
        public async Task<byte[]> GetAsByteArrayAsync(CancellationToken cancellationToken = default) => await this.GetAsByteArrayAsync(true, cancellationToken).ConfigureAwait(false);

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
                this.Encryption.Password = password;
            }

            return await this.GetAsByteArrayAsync(true, cancellationToken).ConfigureAwait(false);
        }

        private async Task ConstructNewFileAsync(string password, CancellationToken cancellationToken)
        {
            MemoryStream? ms = RecyclableMemory.GetStream();
            this._stream ??= RecyclableMemory.GetStream();

            this.File?.Refresh();

            if (this.File != null && this.File.Exists)
            {
                if (password != null)
                {
                    EncryptedPackageHandler? encrHandler = new EncryptedPackageHandler();
                    this.Encryption.IsEncrypted = true;
                    this.Encryption.Password = password;
                    ms.Dispose();
                    ms = encrHandler.DecryptPackage(this.File, this.Encryption);
                }
                else
                {
                    await WriteFileToStreamAsync(this.File.FullName, ms, cancellationToken).ConfigureAwait(false);
                }

                try
                {
                    this._zipPackage = new Packaging.ZipPackage(ms);
                }
                catch (Exception ex)
                {
                    if (password == null && await CompoundDocumentFile.IsCompoundDocumentAsync(this.File, cancellationToken).ConfigureAwait(false))
                    {
                        throw new
                            Exception("Cannot open the package. The package is an OLE compound document. If this is an encrypted package, please supply the password",
                                      ex);
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
                this._zipPackage = new Packaging.ZipPackage(ms);
                ms.Dispose();
                this.CreateBlankWb();
            }
        }

        private static async Task WriteFileToStreamAsync(string path, Stream stream, CancellationToken cancellationToken)
        {
#if NETSTANDARD2_1
            await using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
#else
            using FileStream? fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
#endif
            byte[]? buffer = new byte[4096];
            int read;

            while ((read = await fileStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await stream.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            }
        }

#endif
    }
}