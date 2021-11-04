using System;
using System.IO;

namespace OfficeOpenXml.Utils
{
    internal static class FileHelper
    {
        internal static string GetRelativeFile(FileInfo sourceFile, FileInfo targetFile)
        {
            var sourceDir = sourceFile.DirectoryName ?? "";
            var targetDir = targetFile.DirectoryName ?? ""; ;
            string[] source = sourceDir.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            string[] target = targetDir.Split(new char[]{ '\\' }, StringSplitOptions.RemoveEmptyEntries);

            int slen = source.Length;
            int i = 0;
            while (i < slen && i < target.Length && source[i] == target[i])
            {
                i++;
            }
            if (i == 0) return targetFile.FullName;
            string dirUp = "";
            for (int s = i; s < slen; s++)
            {
                dirUp += "..\\";
            }
            string path = "";
            for (int t = i; t < target.Length; t++)
            {
                path += (path == "" ? "" : "\\") + target[t];
            }
            return dirUp + path +(path == "" ? "" : "\\") + targetFile.Name;
        }

    }
}
