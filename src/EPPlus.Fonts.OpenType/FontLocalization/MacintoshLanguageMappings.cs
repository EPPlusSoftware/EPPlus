/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
  02/27/2026         EPPlus Software AB           Thread-safe initialization
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontLocalization
{
    internal static class MacintoshLanguageMappings
    {
        private static readonly object _syncRoot = new object();
        private static volatile bool _initialized = false;
        private static readonly Dictionary<int, LanguageMapping> _mappings = new Dictionary<int, LanguageMapping>();

        private static void AddMapping(int hexNumber, Languages language)
        {
            // Called only within lock, no additional lock needed
            var mapping = LanguageMapping.Create(hexNumber, language);
            _mappings.Add(mapping.code, mapping);
        }

        private static void CreateMappings()
        {
            AddMapping(0, Languages.English);
            AddMapping(1, Languages.French);
            AddMapping(2, Languages.German);
            AddMapping(3, Languages.Italian);
            AddMapping(4, Languages.Dutch);
            AddMapping(5, Languages.Swedish);
            AddMapping(6, Languages.Spanish);
            AddMapping(7, Languages.Danish);
            AddMapping(8, Languages.Portuguese);
            AddMapping(9, Languages.Norwegian);
            AddMapping(10, Languages.Hebrew);
            AddMapping(11, Languages.Japanese);
            AddMapping(12, Languages.Arabic);
            AddMapping(13, Languages.Finnish);
            AddMapping(14, Languages.Greek);
            AddMapping(15, Languages.Icelandic);
            AddMapping(16, Languages.Maltese);
            AddMapping(17, Languages.Turkish);
            AddMapping(18, Languages.Croatian);
            AddMapping(19, Languages.Chinese_Traditional);
            AddMapping(20, Languages.Urdu);
            AddMapping(21, Languages.Hindi);
            AddMapping(22, Languages.Thai);
            AddMapping(23, Languages.Korean);
            AddMapping(24, Languages.Lithuanian);
            AddMapping(25, Languages.Polish);
            AddMapping(26, Languages.Hungarian);
            AddMapping(27, Languages.Estonian);
            AddMapping(28, Languages.Latvian);
            AddMapping(29, Languages.Sami);
            AddMapping(30, Languages.Faroese);
            AddMapping(31, Languages.Farsi);
            AddMapping(32, Languages.Russian);
            AddMapping(33, Languages.Chinese_Simplified);
            AddMapping(34, Languages.Flemish);
            AddMapping(35, Languages.Irish);
            AddMapping(36, Languages.Albanian);
            AddMapping(37, Languages.Romanian);
            AddMapping(38, Languages.Czech);
            AddMapping(39, Languages.Slovak);
            AddMapping(40, Languages.Slovenian);
            AddMapping(41, Languages.Yiddish);
            AddMapping(42, Languages.Serbian);
            AddMapping(43, Languages.Macedonian);
            AddMapping(44, Languages.Bulgarian);
            AddMapping(45, Languages.Ukrainian);
            AddMapping(46, Languages.Byelorussian);
            AddMapping(47, Languages.Uzbek);
            AddMapping(48, Languages.Kazakh);
            AddMapping(49, Languages.Azeri_Cyrillic);
            AddMapping(50, Languages.Azeri_Arabic);
            AddMapping(51, Languages.Armenian);
            AddMapping(52, Languages.Georgian);
            AddMapping(53, Languages.Moldavian);
            AddMapping(54, Languages.Kirghiz);
            AddMapping(55, Languages.Tajiki);
            AddMapping(56, Languages.Turkmen);
            AddMapping(57, Languages.Mongolian_Traditional);
            AddMapping(58, Languages.Mongolian_Cyrillic);
            AddMapping(59, Languages.Pashto);
            AddMapping(60, Languages.Kurdish);
            AddMapping(61, Languages.Kashmiri);
            AddMapping(62, Languages.Sindhi);
            AddMapping(63, Languages.Tibetan);
            AddMapping(64, Languages.Nepali);
            AddMapping(65, Languages.Sanskrit);
            AddMapping(66, Languages.Marathi);
            AddMapping(67, Languages.Bengali);
            AddMapping(68, Languages.Assamese);
            AddMapping(69, Languages.Gujarati);
            AddMapping(70, Languages.Punjabi);
            AddMapping(71, Languages.Oriya);
            AddMapping(72, Languages.Malayalam);
            AddMapping(73, Languages.Kannada);
            AddMapping(74, Languages.Tamil);
            AddMapping(75, Languages.Telugu);
            AddMapping(76, Languages.Sinhalese);
            AddMapping(77, Languages.Burmese);
            AddMapping(78, Languages.Khmer);
            AddMapping(79, Languages.Lao);
            AddMapping(80, Languages.Vietnamese);
            AddMapping(81, Languages.Indonesian);
            AddMapping(82, Languages.Tagalog);
            AddMapping(83, Languages.Malay_Roman);
            AddMapping(84, Languages.Malay_Arabic);
            AddMapping(85, Languages.Amharic);
            AddMapping(86, Languages.Tigrinya);
            AddMapping(87, Languages.Galla);
            AddMapping(88, Languages.Somali);
            AddMapping(89, Languages.Swahili);
            AddMapping(90, Languages.Kinyarwanda);
            AddMapping(91, Languages.Rundi);
            AddMapping(128, Languages.Welsh);
            AddMapping(129, Languages.Basque);
            AddMapping(130, Languages.Catalan);
            AddMapping(131, Languages.Latin);
            AddMapping(132, Languages.Quechua);
            AddMapping(133, Languages.Gurani);
            AddMapping(134, Languages.Aymara);
            AddMapping(135, Languages.Tatar);
            AddMapping(136, Languages.Uighur);
            AddMapping(137, Languages.Dzongkha);
            AddMapping(138, Languages.Javanese_Roman);
            AddMapping(139, Languages.Sundanese_Roman);
            AddMapping(140, Languages.Galician);
            AddMapping(141, Languages.Afrikaans);
            AddMapping(142, Languages.Breton);
            AddMapping(143, Languages.Inuktitut);
            AddMapping(144, Languages.Galician);
            AddMapping(145, Languages.Galician);
            AddMapping(146, Languages.Irish);
            AddMapping(147, Languages.Tongan);
            AddMapping(148, Languages.Greek_Polytonic);
            AddMapping(149, Languages.Greenlandic);
            AddMapping(150, Languages.Azeri_Latin);
        }

        public static IDictionary<int, LanguageMapping> Mappings
        {
            get
            {
                // Double-checked locking pattern, compatible with .NET 3.5+
                if (!_initialized)
                {
                    lock (_syncRoot)
                    {
                        if (!_initialized)
                        {
                            CreateMappings();
                            _initialized = true;
                        }
                    }
                }
                return _mappings;
            }
        }
    }
}