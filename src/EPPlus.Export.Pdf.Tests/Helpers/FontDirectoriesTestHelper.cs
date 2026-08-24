/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/21/2025         EPPlus Software AB           Test infrastructure
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Export.Pdf.Tests.Helpers
{
    public class FontDirectoriesTestHelper
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();
        private static string _testOutputPath = @"c:\epplusTest\Fonts\";
        private static bool _testOutputAvailable = false;
        private static bool _initialized = false;
        private static object _syncRoot = new object();

        /// <summary>
        /// Gets the font folder path (for reading test fonts)
        /// </summary>
        public static string FontFolder => _fontFolder;

        /// <summary>
        /// Gets the list of font folder paths (for reading test fonts)
        /// </summary>
        public static List<string> FontFolders => new List<string>(_fontFolders);

        /// <summary>
        /// Gets the test output path for saving subset fonts (c:\epplusTest\Fonts\)
        /// Returns null if path is not available (e.g. in CI/CD)
        /// </summary>
        public static string TestOutputPath => _testOutputAvailable ? _testOutputPath : null;

        /// <summary>
        /// Gets whether test output path is available (false in CI/CD environments)
        /// </summary>
        public static bool IsTestOutputAvailable => _testOutputAvailable;

        /// <summary>
        /// Initializes test environment. Thread-safe singleton.
        /// Call this from [ClassInitialize] in each test class.
        /// </summary>
        /// <param name="testContext">MSTest context (unused but required by MSTest)</param>
        public static void ClassInitialize(TestContext testContext)
        {
            if (!_initialized)
            {
                lock (_syncRoot)
                {
                    if (!_initialized)
                    {
                        // Input folder (test fonts from project) - ALWAYS AVAILABLE
                        _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
                        _fontFolders = new List<string> { _fontFolder };

                        // Output folder (c:\epplusTest\Fonts\) - OPTIONAL
                        // Check if c:\epplusTest exists (developer machines only)
                        var epplusTestRoot = @"c:\epplusTest";
                        if (Directory.Exists(epplusTestRoot))
                        {
                            _testOutputAvailable = true;
                            TryCreateTestOutputPath();
                        }
                        else
                        {
                            _testOutputAvailable = false;
                        }

                        OpenTypeFonts.ClearFontCache();
                        _initialized = true;
                    }
                }
            }
        }

        /// <summary>
        /// Saves a font to the test output folder.
        /// SKIPS silently if test output path is not available (CI/CD).
        /// If the file exists, it will be deleted first.
        /// </summary>
        /// <param name="font">Font to save</param>
        /// <param name="fileName">Filename (e.g. "subset_Roboto_abc.ttf")</param>
        /// <returns>FileInfo for the saved file, or null if output not available</returns>
        public static FileInfo SaveFontToOutput(OpenTypeFont font, string fileName)
        {
            if (!_testOutputAvailable)
            {
                // Silently skip in CI/CD - this is for manual inspection only
                return null;
            }

            if (font == null)
                throw new ArgumentNullException(nameof(font));

            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("Filename cannot be empty", nameof(fileName));

            TryCreateTestOutputPath();

            var fi = new FileInfo(_testOutputPath + fileName);
            if (fi.Exists)
            {
                fi.Delete();
            }

            var bytes = font.Serialize();
            File.WriteAllBytes(fi.FullName, bytes);

            return fi;
        }

        /// <summary>
        /// Gets a FileInfo for an output file, creating subdirectories if needed.
        /// Returns null if test output path is not available (CI/CD).
        /// </summary>
        /// <param name="subPath">Subdirectory under c:\epplusTest\Fonts\ (e.g. "Roboto")</param>
        /// <param name="fileName">Filename (e.g. "subset_abc.ttf")</param>
        /// <returns>FileInfo for the output file, or null if output not available</returns>
        public static FileInfo GetOutputFile(string subPath, string fileName)
        {
            if (!_testOutputAvailable)
            {
                return null;
            }

            var path = _testOutputPath + subPath;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }

            return new FileInfo(path + fileName);
        }

        /// <summary>
        /// Tries to create the test output path. Silent fail if not possible.
        /// </summary>
        private static void TryCreateTestOutputPath()
        {
            if (!_testOutputAvailable)
                return;

            try
            {
                if (!Directory.Exists(_testOutputPath))
                {
                    Directory.CreateDirectory(_testOutputPath);
                }
            }
            catch
            {
                // Silent fail - output is optional
                _testOutputAvailable = false;
            }
        }

        /// <summary>
        /// Checks if a font file exists in the output folder.
        /// Returns false if test output path is not available.
        /// </summary>
        /// <param name="fileName">Filename to check</param>
        /// <returns>True if file exists</returns>
        public static bool ExistsOutputFont(string fileName)
        {
            if (!_testOutputAvailable)
                return false;

            var fi = new FileInfo(_testOutputPath + fileName);
            return fi.Exists;
        }

        /// <summary>
        /// Deletes a font file from the output folder if it exists.
        /// Skips silently if test output path is not available.
        /// </summary>
        /// <param name="fileName">Filename to delete</param>
        public static void DeleteOutputFont(string fileName)
        {
            if (!_testOutputAvailable)
                return;

            var fi = new FileInfo(_testOutputPath + fileName);
            if (fi.Exists)
            {
                fi.Delete();
            }
        }
    }
}