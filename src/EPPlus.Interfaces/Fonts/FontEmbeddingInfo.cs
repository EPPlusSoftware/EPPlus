namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Information passed to the <see cref="IEpplusFontConfiguration.OnFontEmbedding"/>
    /// callback so the caller can decide how a font should be embedded.
    /// </summary>
    public class FontEmbeddingInfo
    {
        public FontEmbeddingInfo(string fontName, FontEmbeddingRestriction restriction)
            : this(fontName, restriction, FontEmbeddingPermission.Installable, false, FontRenderTarget.Document)
        {
        }

        public FontEmbeddingInfo(
            string fontName,
            FontEmbeddingRestriction restriction,
            FontEmbeddingPermission permission,
            bool bitmapEmbeddingOnly,
            FontRenderTarget target)
            : this(fontName, restriction, permission, bitmapEmbeddingOnly, target, false, null)
        {
        }

        public FontEmbeddingInfo(
            string fontName,
            FontEmbeddingRestriction restriction,
            FontEmbeddingPermission permission,
            bool bitmapEmbeddingOnly,
            FontRenderTarget target,
            bool hasOpenLicense,
            string licenseDescription)
        {
            FontName = fontName;
            Restriction = restriction;
            Permission = permission;
            BitmapEmbeddingOnly = bitmapEmbeddingOnly;
            Target = target;
            HasOpenLicense = hasOpenLicense;
            LicenseDescription = licenseDescription;
        }

        /// <summary>The full name of the font being prepared for embedding.</summary>
        public string FontName { get; private set; }

        /// <summary>The restriction the font declares via its OS/2 fsType field.</summary>
        public FontEmbeddingRestriction Restriction { get; private set; }

        /// <summary>
        /// The usage permission the font declares via its OS/2 fsType field. Distinguishes
        /// Preview &amp; Print from Editable and Installable, which <see cref="Restriction"/> does not.
        /// </summary>
        public FontEmbeddingPermission Permission { get; private set; }

        /// <summary>
        /// True when the font's OS/2 fsType field declares that only bitmap data may be embedded.
        /// The font's outlines must then not be embedded.
        /// </summary>
        public bool BitmapEmbeddingOnly { get; private set; }

        /// <summary>
        /// The output the font is embedded in. <see cref="FontRenderTarget.Document"/> is a PDF;
        /// <see cref="FontRenderTarget.Web"/> is an SVG or HTML file, where an embedded font is
        /// delivered as a web font. Font licences often treat the two differently.
        /// </summary>
        public FontRenderTarget Target { get; private set; }

        /// <summary>
        /// True when the font's name table (licence description, nameID 13, or licence URL,
        /// nameID 14) names an open licence that EPPlus recognises: SIL Open Font License,
        /// Apache License, Ubuntu Font Licence or Bitstream Vera. The default web policy embeds
        /// only such fonts. False for commercial fonts and for fonts that carry no licence text.
        /// </summary>
        public bool HasOpenLicense { get; private set; }

        /// <summary>
        /// The font's licence description (nameID 13) as written in the font, or null when it has
        /// none. Lets the callback apply its own rules, for instance for a licence EPPlus does not
        /// recognise.
        /// </summary>
        public string LicenseDescription { get; private set; }
    }
}