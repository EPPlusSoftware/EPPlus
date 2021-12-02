/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// The base class for a theme
    /// </summary>
    public class ExcelThemeBase : XmlHelper, IPictureRelationDocument
    {
        readonly string _colorSchemePath = "{0}a:clrScheme";
        readonly string _fontSchemePath = "{0}a:fontScheme";
        readonly string _fmtSchemePath = "{0}a:fmtScheme";
        readonly ExcelPackage _pck;
        Dictionary<string, HashInfo> _hashes=new Dictionary<string, HashInfo>();
        internal ExcelThemeBase(ExcelPackage package, XmlNamespaceManager nsm, ZipPackageRelationship rel, string path)
            : base(nsm, null)
        {
            ThemeUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            Part = package.ZipPackage.GetPart(ThemeUri);
            RelationshipId = rel.Id;
            ThemeXml = new XmlDocument();
            LoadXmlSafe(ThemeXml, Part.GetStream());
            TopNode = ThemeXml.DocumentElement;

            _colorSchemePath = string.Format(_colorSchemePath, path);
            _fontSchemePath = string.Format(_fontSchemePath, path);
            _fmtSchemePath = string.Format(_fmtSchemePath, path);
            _pck = package;
            if (!NameSpaceManager.HasNamespace("a")) NameSpaceManager.AddNamespace("a", ExcelPackage.schemaDrawings);
        }
        internal Uri ThemeUri { get; set; }
        internal ZipPackagePart Part { get; set; }
        /// <summary>
        /// The Theme Xml
        /// </summary>
        public XmlDocument ThemeXml { get; internal set; }
        internal string RelationshipId { get; set; }
        internal ExcelColorScheme _colorScheme = null;
        /// <summary>
        /// Defines the color scheme
        /// </summary>
        public ExcelColorScheme ColorScheme
        {
            get
            {
                if (_colorScheme == null)
                {
                    _colorScheme = new ExcelColorScheme(NameSpaceManager, TopNode.SelectSingleNode(_colorSchemePath, NameSpaceManager));
                }
                return _colorScheme;
            }
        }
        internal ExcelFontScheme _fontScheme = null;
        /// <summary>
        /// Defines the font scheme
        /// </summary>
        public ExcelFontScheme FontScheme
        {
            get
            {
                if (_fontScheme == null)
                {
                    _fontScheme = new ExcelFontScheme(_pck,NameSpaceManager, TopNode.SelectSingleNode(_fontSchemePath, NameSpaceManager));
                }
                return _fontScheme;
            }
        }
        private ExcelFormatScheme _formatScheme = null;
        /// <summary>
        /// The background fill styles, effect styles, fill styles, and line styles which define the style matrix for a theme
        /// </summary>
        public ExcelFormatScheme FormatScheme
        {
            get
            {
                if (_formatScheme == null)
                {
                    _formatScheme = new ExcelFormatScheme(NameSpaceManager, TopNode.SelectSingleNode(_fmtSchemePath, NameSpaceManager), this);
                }
                return _formatScheme;
            }
        }

        ExcelPackage IPictureRelationDocument.Package { get => _pck; }

        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes { get => _hashes; }

        ZipPackagePart IPictureRelationDocument.RelatedPart { get => Part; }

        Uri IPictureRelationDocument.RelatedUri { get => ThemeUri; }
    }
}
