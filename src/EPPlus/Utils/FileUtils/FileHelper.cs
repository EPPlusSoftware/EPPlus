using System;
using System.IO;

namespace OfficeOpenXml.Utils.FileUtils
{
    internal static class FileHelper
    {
        internal static string GetRelativeFile(FileInfo sourceFile, FileInfo targetFile, bool addFileProtocolIfAbsolute = false)
        {
            if (sourceFile == null) return (addFileProtocolIfAbsolute ? "file:///" : "") + targetFile.FullName;
            var sourceDir = sourceFile.DirectoryName ?? "";
            var targetDir = targetFile.DirectoryName ?? "";
            string[] source = sourceDir.Split(new char[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
            string[] target = targetDir.Split(new char[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

            int slen = source.Length;
            int i = 0;
            while (i < slen && i < target.Length && source[i] == target[i])
            {
                i++;
            }
            if (i == 0)
            {
                if (addFileProtocolIfAbsolute && targetFile.FullName.StartsWith("file:///") == false)
                {
                    return "file:///" + targetFile.FullName;
                }
                else
                {
                    return targetFile.FullName;
                }
            }

            string dirUp = "";
            for (int s = i; s < slen; s++)
            {
                dirUp += $"..{Path.DirectorySeparatorChar}";
            }
            string path = "";
            for (int t = i; t < target.Length; t++)
            {
                path += (path == "" ? "" : $"{Path.DirectorySeparatorChar}") + target[t];
            }
            return dirUp + path + (path == "" ? "" : $"{Path.DirectorySeparatorChar}") + targetFile.Name;
        }

        internal static bool IsFileNameValid(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Validates a path to a file that is about to be created or overwritten, so that errors can be
        /// reported before any expensive work is done. Does not touch the file system beyond checking
        /// that the directory exists and that the path does not refer to an existing directory.
        /// Access rights and locked files are not checked; they surface when the file is opened.
        /// </summary>
        /// <param name="filePath">The path to validate. Relative paths are resolved against the current directory.</param>
        /// <param name="paramName">The parameter name to report in exceptions.</param>
        /// <returns>The full path of the file.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="filePath"/> is null.</exception>
        /// <exception cref="ArgumentException">The path is empty, has an invalid format, contains invalid characters, has no file name or refers to a directory.</exception>
        /// <exception cref="PathTooLongException">The path exceeds the system-defined maximum length.</exception>
        /// <exception cref="DirectoryNotFoundException">The directory of the file does not exist.</exception>
        internal static string ValidateOutputFilePath(string filePath, string paramName)
        {
            if (filePath == null)
                throw new ArgumentNullException(paramName);
            if (filePath.Trim().Length == 0)
                throw new ArgumentException("The file path cannot be empty or consist only of white space.", paramName);

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(filePath);
            }
            catch (ArgumentException ex)
            {
                throw new ArgumentException($"The file path '{filePath}' is not valid.", paramName, ex);
            }
            catch (NotSupportedException ex)
            {
                throw new ArgumentException($"The file path '{filePath}' has an invalid format.", paramName, ex);
            }

            if (Directory.Exists(fullPath))
                throw new ArgumentException($"The file path '{filePath}' refers to a directory.", paramName);

            // Also catches a path ending with a directory separator, which gives an empty file name.
            if (!IsFileNameValid(Path.GetFileName(fullPath)))
                throw new ArgumentException($"The file path '{filePath}' does not contain a valid file name.", paramName);

            var directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                throw new DirectoryNotFoundException($"Could not find the directory '{directory}'.");

            return fullPath;
        }
    }
}