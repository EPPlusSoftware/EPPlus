/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System;
using System.IO;

namespace EPPlus.Export.Pdf.Utils.Platform
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
