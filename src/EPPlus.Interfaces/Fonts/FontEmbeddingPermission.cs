namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// The usage permission a font declares via bits 0-3 of its OS/2 fsType field.
    /// Pure interpretation of fsType — carries no policy about what EPPlus does.
    /// </summary>
    public enum FontEmbeddingPermission
    {
        /// <summary>No usage bit set. The font may be embedded and permanently installed by the recipient.</summary>
        Installable,
        /// <summary>The font must not be embedded (Restricted License).</summary>
        RestrictedLicense,
        /// <summary>The font may be embedded in documents that are viewed or printed, but not edited.</summary>
        PreviewAndPrint,
        /// <summary>The font may be embedded in documents that the recipient may view, print and edit.</summary>
        Editable,
    }
}