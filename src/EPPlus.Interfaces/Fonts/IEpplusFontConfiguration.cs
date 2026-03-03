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
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Global font configuration for EPPlus.
    /// Accessed exclusively via <c>OpenTypeFonts.Configure(Action&lt;IEpplusFontConfiguration&gt;)</c>.
    /// Changes are global and persist for the lifetime of the application unless
    /// <see cref="Reset"/> is called.
    /// </summary>
    public interface IEpplusFontConfiguration
    {
        /// <summary>
        /// Adds a font-name fallback chain.
        /// "If Arial is not found, try Helvetica, then Roboto."
        /// Additive — does not replace previously added fallbacks for the same font.
        /// </summary>
        /// <param name="primaryFont">The font that may be missing.</param>
        /// <param name="fallbacks">
        /// One or more font names to try in order when <paramref name="primaryFont"/>
        /// cannot be resolved.
        /// </param>
        /// <returns>This instance, for fluent chaining.</returns>
        IEpplusFontConfiguration AddFallback(string primaryFont, params string[] fallbacks);

        /// <summary>
        /// Replaces the font resolver entirely.
        /// The user takes full responsibility for font resolution.
        /// EPPlus built-in fallbacks (Archivo Narrow) will NOT be applied.
        /// </summary>
        /// <param name="resolver">
        /// A custom <see cref="IFontResolver"/> that returns raw TTF/OTF bytes.
        /// </param>
        /// <returns>This instance, for fluent chaining.</returns>
        IEpplusFontConfiguration SetFontResolver(IFontResolver resolver);

        /// <summary>
        /// Restores all settings to factory defaults:
        /// <list type="bullet">
        ///   <item>Clears all font-name fallbacks.</item>
        ///   <item>Restores <c>DefaultFontResolver</c> (with Archivo Narrow built-in fallback).</item>
        ///   <item>Restores <c>DefaultFontProvider</c> (with Noto Emoji + Noto Sans Math chain).</item>
        ///   <item>Clears the font cache.</item>
        /// </list>
        /// </summary>
        /// <returns>This instance, for fluent chaining.</returns>
        IEpplusFontConfiguration Reset();
    }
}