using System;
using System.IO;

namespace LevitJames.Core
{
    /// <summary>
    ///     A class for managing uniquely named Temporary files in the users temporary directory.
    /// </summary>
    public sealed class TemporaryFile : IDisposable
    {
#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif


        /// <summary>
        ///     Creates a new TemporaryFile instance including the new temporary file.
        /// </summary>
        public TemporaryFile() : this(createTempFile: true) { }

        /// <summary>
        ///     Creates a new TemporaryFile instance, optionally creating the temporary file
        /// </summary>
        /// <param name="createTempFile">true to create the temporary file false otherwise.</param>
        public TemporaryFile(bool createTempFile)
        {
            if (createTempFile)
            {
                FileName = Path.GetTempFileName();
            }
            else
                FileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        /// <summary>
        ///     Create a new TemporaryFile Instance and the temporary file then copies the contents of the file provided by the
        ///     fileName over to the new temporary file.
        /// </summary>
        /// <param name="baseFile">The file who's content should copied to the temporary file</param>
        public TemporaryFile(string baseFile) : this(baseFile, keepExtension: false) { }

        /// <summary>
        ///     Create a new TemporaryFile Instance and the temporary file then copies the contents of the file provided by the
        ///     fileName over to the new temporary file.
        ///     Optionally keeps the extension from the base file.
        /// </summary>
        /// <param name="baseFile">The file who's content should copied to the temporary file</param>
        /// <param name="keepExtension">If true, keeps the same extension for the temporary file</param>
        public TemporaryFile(string baseFile, bool keepExtension) : this()
        {
            if (keepExtension)
            {
                FileName = Path.ChangeExtension(FileName, Path.GetExtension(baseFile));
            }

            CopyFrom(baseFile);
        }


        /// <summary>
        ///     Creates an empty file in the specified path provided.
        /// </summary>
        /// <param name="path">The path to create the file in.</param>
        /// <returns>If the path does not exist an exception is thrown.</returns>
        public static TemporaryFile FromPath(string path)
        {
            var tempFile = new TemporaryFile { FileName = Path.Combine(path, Path.GetRandomFileName()) };
            File.Create(tempFile.FileName).Dispose();
            return tempFile;
        }

        /// <summary>
        ///     Creates an empty file in the specified path provided.
        /// </summary>
        /// <param name="path">The path to create the file in.</param>
        /// <param name="fileExtension">The file extension for the fileName.</param>
        /// <returns>If the path does not exist an exception is thrown.</returns>
        public static TemporaryFile FromPath(string path, string fileExtension)
        {

            var resolvedFileExtension = Path.GetExtension(fileExtension);
            if (string.IsNullOrEmpty(resolvedFileExtension))
                resolvedFileExtension = '.' + fileExtension;

            var tempFile = new TemporaryFile();
            try
            {
                do
                {
                    //Don't use GetTempFileName as it created the file. GetRandomFileName does not.
                    //Note: GetRandomFileName is just the file name and contains no path.
                    tempFile.FileName = Path.ChangeExtension(Path.Combine(path, Path.GetRandomFileName()), resolvedFileExtension);
                } while (File.Exists(tempFile.FileName));

                //Finally create an empty file so this method is consistent with Path.GetTempFileName() which created the file.
                File.Create(tempFile.FileName).Dispose();

                return tempFile;
            }
            catch
            {
                tempFile.Dispose();
                throw;
            }
        }


        /// <summary>
        ///     Creates a new TemporaryFile instance using the supplied file extension.
        /// </summary>
        /// <param name="fileExtension">The file extension to append to the temporary file.</param>
        /// <returns>A new TemporaryFile instance</returns>
        public static TemporaryFile FromFileExtension(string fileExtension) => FromPath(Path.GetTempPath(), fileExtension);

        /// <summary>
        ///     Creates a new TemporaryFile instance with a specific file extension.
        /// </summary>
        /// <param name="fileExtension"></param>
        /// <param name="baseFile">The file who's content should copied to the temporary file</param>
        public static TemporaryFile FromFileExtension(string fileExtension, string baseFile)
        {
            var tf = FromFileExtension(fileExtension);
            try
            {
                tf.CopyFrom(baseFile);
            }
            catch
            {
                tf?.Dispose();
                throw;
            }

            return tf;
        }


        /// <summary>
        ///     Returns sets if the file should be deleted when this class is disposed.
        /// </summary>
        /// <remarks>Call this if the file need to be kept open and or owned by other routines.</remarks>
        public bool KeepFileOnDispose { get; set; }


        /// <summary>
        ///     Returns file name of the temporary file.
        /// </summary>
        public string FileName { get; private set; }


        /// <summary>
        ///     Deletes the temporary file.
        /// </summary>
        public void Delete()
        {
            Delete(true);
        }

        private void Delete(bool throwOnFail)
        {
            if (!File.Exists(FileName))
                return;

            try
            {
                try
                {
                    File.Delete(FileName);
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(250);
                    File.Delete(FileName);
                }
                catch (UnauthorizedAccessException)
                {
                    System.Threading.Thread.Sleep(250);
                    File.Delete(FileName);
                }
            }
            catch (Exception)
            {
                if (throwOnFail)
                    throw;
            }
        }

        /// <summary>
        ///     Writes all the text to the temporary file.
        /// </summary>
        /// <param name="contents">The contents to write to the temporary file.</param>
        public void Write(string contents)
        {
            File.WriteAllText(FileName, contents);
        }

        /// <summary>
        ///     Writes all the bytes to the temporary file.
        /// </summary>
        /// <param name="contents">The contents to write to the temporary file.</param>
        public void Write(byte[] contents)
        {
            File.WriteAllBytes(FileName, contents);
        }


        /// <summary>
        ///     Writes the stream out to the temporary file.
        /// </summary>
        /// <param name="contents">The contents to write to the temporary file.</param>
        public void Write(Stream contents)
        {
            var sr = new BinaryReader(contents);
            File.WriteAllBytes(FileName, sr.ReadBytes((int)contents.Length));
        }

        /// <summary>
        ///     Writes all the lines to the temporary file.
        /// </summary>
        /// <param name="contents">The contents to write to the temporary file.</param>
        public void Write(string[] contents)
        {
            File.WriteAllLines(FileName, contents);
        }


        /// <summary>
        ///     Reads all the bytes to the temporary file.
        /// </summary>
        public byte[] ReadAllBytes()
        {
            return File.Exists(FileName) ? File.ReadAllBytes(FileName) : new byte[] { };
        }


        /// <summary>
        ///     Reads all the bytes from the temporary file into the supplied Stream.
        /// </summary>
        /// <param name="stream">The stream to store the contents of the temporary file in.</param>
        public void ReadToStream(Stream stream)
        {
            if (!File.Exists(FileName))
                return;

            var sw = new BinaryWriter(stream);
            sw.Write(File.ReadAllBytes(FileName));
        }


        /// <summary>
        ///     Returns an array of strings containing all the lines in the temporary file.
        /// </summary>
        public string[] ReadAllLines()
        {
            return File.Exists(FileName) ? File.ReadAllLines(FileName) : new string[] { };
        }


        /// <summary>
        ///     Reads all the bytes to the temporary file
        /// </summary>
        public string ReadAllText()
        {
            return File.Exists(FileName) ? File.ReadAllText(FileName) : string.Empty;
        }


        /// <summary>
        ///     The name of the original file used to Copy the temporary file from.
        /// </summary>
        public string OriginalFileName { get; private set; }


        /// <summary>
        ///     Copies the supplied file to the temporary file, overwriting the temporary file if it exists.
        /// </summary>
        /// <param name="fileToCopy">The path of the file to copy.</param>
        public void CopyFrom(string fileToCopy)
        {
            if (File.Exists(fileToCopy))
            {
                File.Copy(fileToCopy, FileName, overwrite: true);
                OriginalFileName = fileToCopy;
            }
        }

        /// <summary>
        ///     Copies a file to the temporary file
        /// </summary>
        /// <param name="copyToFile"></param>
        public void CopyTo(string copyToFile)
        {
            File.Copy(FileName, copyToFile, overwrite: true);
        }


        /// <summary>
        ///     Finalizer for TemporaryFile
        /// </summary>
        ~TemporaryFile()
        {
#if (TRACK_DISPOSED)
                LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(disposing: false);
        }

        /// <summary>
        ///     If KeepFileOnDispose is false then the temporary file is deleted; otherwise control of deleting the file is passed
        ///     to the calling code.
        ///     The FileName property is also set to null;
        /// </summary>
        public void Dispose() => Dispose(disposing: true);

        private void Dispose(bool disposing)
        {
            if (KeepFileOnDispose == false)
            {
                Delete(throwOnFail: disposing); //Don't throw if called from finalizer
            }

            FileName = null;
            if (disposing)
                GC.SuppressFinalize(this);
        }
    }
}
