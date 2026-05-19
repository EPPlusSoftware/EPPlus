/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/27/2026         EPPlus Software AB           Initial implementation
  05/06/2026         EPPlus Software AB           Property-based transactional configuration
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Global font configuration for EPPlus.
    /// Accessed exclusively via <c>OpenTypeFonts.Configure(Action&lt;IEpplusFontConfiguration&gt;)</c>.
    /// Changes made inside a Configure callback are applied as a single transaction —
    /// when the callback returns, the font resolver is rebuilt and all font caches are cleared.
    /// </summary>
    public interface IEpplusFontConfiguration
    {
        /// <summary>
        /// Additional directories to search for font files, beyond the system font directories.
        /// Mutate this list inside a Configure callback to add or remove search paths.
        /// </summary>
        IList<string> FontDirectories { get; }

        /// <summary>
        /// Whether the operating system's standard font directories should be searched.
        /// Defaults to <c>true</c>. Set to <c>false</c> to restrict font resolution to the
        /// directories listed in <see cref="FontDirectories"/>.
        /// </summary>
        bool SearchSystemDirectories { get; set; }

        /// <summary>
        /// The font resolver responsible for producing raw TTF/OTF bytes for a requested font.
        /// Set this to replace the default resolver entirely. When a custom resolver is set,
        /// the EPPlus built-in fallback chains and Archivo Narrow ultimate fallback are bypassed —
        /// the resolver is fully responsible for handling missing fonts.
        /// </summary>
        IFontResolver FontResolver { get; set; }

        /// <summary>
        /// User-defined font-name fallback chains.
        /// Each entry maps a font name to an ordered list of fallbacks to try when the primary
        /// is unavailable. Mutate this dictionary inside a Configure callback to add or remove
        /// chains. User chains are tried before the EPPlus built-in fallback chains.
        /// </summary>
        IDictionary<string, string[]> FontFallbacks { get; }

        /// <summary>
        /// Restores all settings to factory defaults:
        /// <list type="bullet">
        ///   <item>Clears <see cref="FontDirectories"/>.</item>
        ///   <item>Sets <see cref="SearchSystemDirectories"/> to <c>true</c>.</item>
        ///   <item>Restores the default <see cref="FontResolver"/> (with Archivo Narrow built-in fallback).</item>
        ///   <item>Clears <see cref="FontFallbacks"/>.</item>
        /// </list>
        /// </summary>
        void Reset();
    }
}