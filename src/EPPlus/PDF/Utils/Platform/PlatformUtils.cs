using System;
using System.IO;

namespace OfficeOpenXml.PDF.Utils.Platform
{
    internal class PlatformUtils
    {
        public enum OperatingSystem
        {
            Windows,
            Mac,
            Linux,
            Unknown
        }
        public static OperatingSystem GetPlatform()
        {
            PlatformID platform = Environment.OSVersion.Platform;

            if (platform == PlatformID.Win32NT || platform == PlatformID.Win32Windows ||
                platform == PlatformID.Win32S || platform == PlatformID.WinCE)
            {
                return OperatingSystem.Windows;
            }

            if (platform == PlatformID.MacOSX)
            {
                return OperatingSystem.Mac;
            }

            if (platform == PlatformID.Unix)
            {
                // macOS has this folder; Linux doesn't
                if (Directory.Exists("/System/Library/CoreServices"))
                    return OperatingSystem.Mac;

                return OperatingSystem.Linux;
            }

            return OperatingSystem.Unknown;
        }
    }
}
