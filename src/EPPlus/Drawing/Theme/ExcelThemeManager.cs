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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// Handels themes in a package
    /// </summary>
    public class ExcelThemeManager
    {
        ExcelWorkbook _wb;
        internal static string _defaultTheme="";
        internal ExcelThemeManager(ExcelWorkbook wb)
        {
            _wb = wb;
        }
        ExcelTheme _theme = null;
        /// <summary>
        /// The current theme. Null if not theme exists.
        /// <seealso cref="CreateDefaultTheme"/>
        /// <seealso cref="Load(FileInfo)"/>
        /// <seealso cref="Load(Stream)"/>
        /// <seealso cref="Load(XmlDocument)"/>
        /// </summary>
        public ExcelTheme CurrentTheme
        {
            get
            {
                if(_theme==null)
                {
                    var rels = _wb.Part.GetRelationshipsByType(ExcelPackage.schemaThemeRelationships);
                    if (rels.Count>0)
                    {                        
                        _theme = new ExcelTheme(_wb, rels.First());
                    }
                }
                return _theme;
            }
        }
        /// <summary>
        /// Create the default theme.
        /// </summary>
        public void CreateDefaultTheme()
        {
            if (CurrentTheme != null)
            {
                throw (new InvalidOperationException("Can't create theme. Theme already exists"));
            }

            if(string.IsNullOrEmpty(_defaultTheme))
            {
                _defaultTheme = StyleResourceManager.GetItem("DefaultTheme.Xml");
            }
            var themeXml = new XmlDocument();   
            themeXml.LoadXml(_defaultTheme);
            Load(themeXml);
        }
        /// <summary>
        /// Delete the current theme
        /// </summary>
        public void DeleteCurrentTheme()
        {
            if(CurrentTheme==null)
            {
                return;
            }
            _wb._package.ZipPackage.DeleteRelationship(_theme.RelationshipId);
            _wb._package.ZipPackage.DeletePart(_theme.ThemeUri);
            _theme = null;
        }
        /// <summary>
        /// Loads a .thmx file, exported from a Spread Sheet Application like Excel
        /// </summary>
        /// <param name="thmxFile">The path to the thmx file</param>
        public void Load(FileInfo thmxFile)
        {
            if(!thmxFile.Exists)
            {
                throw (new FileNotFoundException($"{thmxFile.FullName} does not exist"));
            }

            using (var ms = RecyclableMemory.GetStream(File.ReadAllBytes(thmxFile.FullName)))
            {
                Load(ms);
            }
        }
        /// <summary>
        /// Loads a theme XmlDocument. 
        /// Overwrites any previously set theme settings.
        /// </summary>
        /// <param name="themeXml">The theme xml</param>
        public void Load(XmlDocument themeXml)
        {
            DeleteCurrentTheme();
            if (CurrentTheme == null)
            {
                var uri = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
                var part = _wb._package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeTheme);
                themeXml.Save(part.GetStream());
                var rel = _wb.Part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaThemeRelationships);
                _theme = new ExcelTheme(_wb, rel);
            }
        }
        /// <summary>
        /// Loads a .thmx file as a stream. Thmx files are exported from a Spread Sheet Application like Excel
        /// </summary>
        /// <param name="thmxStream">The thmx file as a stream</param>
        public void Load(Stream thmxStream)
        {
            
            ZipPackage p = new ZipPackage(thmxStream);
            
            var themeManagerRel = p.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").FirstOrDefault();
            if (themeManagerRel != null)
            {
                var themeManager = p.GetPart(themeManagerRel.TargetUri);
                var themeRel = themeManager.GetRelationshipsByType(ExcelPackage.schemaThemeRelationships).FirstOrDefault();
                if (themeRel != null)
                {
                    var themePart = p.GetPart(UriHelper.ResolvePartUri(themeRel.SourceUri, themeRel.TargetUri));
                    var themeXml = new XmlDocument();
                    XmlHelper.LoadXmlSafe(themeXml, themePart.GetStream());
                    Load(themeXml);
                    foreach (var rel in themePart.GetRelationships())
                    {   
                        var partToCopy = p.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                        var uri = UriHelper.ResolvePartUri(_theme.ThemeUri, rel.TargetUri);
                        var part = _wb._package.ZipPackage.CreatePart(uri, partToCopy.ContentType);
                        var stream = part.GetStream();
                        var b = ((MemoryStream)partToCopy.GetStream()).ToArray();
                        stream.Write(b, 0, b.Length);
                        stream.Flush();
                        _theme.Part.CreateRelationship(uri, TargetMode.Internal, rel.RelationshipType);
                    }
                    SetNormalStyle();
                }
                else
                {
                    throw new InvalidDataException("Thmx file is corrupt. Can't find theme part");
                }
            }
            else
            {
                throw new InvalidDataException("Thmx file is corrupt.");
            }
        }

        private void SetNormalStyle()
        {
            if (_wb.Styles.NamedStyles.Count == 0) return;
            var style = GetNormalStyle();
            foreach(var xfs in _wb.Styles.CellXfs)
            {
                if (xfs.XfId == style.StyleXfId)
                {
                    var font = _wb.Styles.Fonts[xfs.FontId];
                    font.Name = CurrentTheme.FontScheme.MinorFont[0].Typeface;
                    font.Family = 2;
                    font.Color.Theme = eThemeSchemeColor.Text1;
                    font.Scheme = "minor";
                }
            }
        }

        private ExcelNamedStyleXml GetNormalStyle()
        {
            return _wb.Styles.GetNormalStyle();
        }
        internal void Save()
        {
            if (CurrentTheme != null)
            {
                _wb._package.SavePart(CurrentTheme.ThemeUri, CurrentTheme.ThemeXml);
            }
        }
    }
}

