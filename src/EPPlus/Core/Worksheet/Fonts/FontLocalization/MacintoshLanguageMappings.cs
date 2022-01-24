using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.FontLocalization
{
    public static class MacintoshLanguageMappings
    {
        private static IDictionary<int, LanguageMapping> _mappings = new Dictionary<int, LanguageMapping>();

        private static void AddMappping(int hexNumber, Languages language)
        {
            var mapping = LanguageMapping.Create(hexNumber, language);
            _mappings.Add(mapping.code, mapping);
        }
        private static void CreateMappings()
        {
            AddMappping(0, Languages.English);
            AddMappping(1, Languages.French);
            AddMappping(2, Languages.German);
            AddMappping(3, Languages.Italian);
            AddMappping(4, Languages.Dutch);
            AddMappping(5, Languages.Swedish);
            AddMappping(6, Languages.Spanish);
            AddMappping(7, Languages.Danish);
            AddMappping(8, Languages.Portuguese);
            AddMappping(9, Languages.Norwegian);
            AddMappping(10, Languages.Hebrew);
            AddMappping(11, Languages.Japanese);
            AddMappping(12, Languages.Arabic);
            AddMappping(13, Languages.Finnish);
            AddMappping(14, Languages.Greek);
            AddMappping(15, Languages.Icelandic);
            AddMappping(16, Languages.Maltese);
            AddMappping(17, Languages.Turkish);
            AddMappping(18, Languages.Croatian);
            AddMappping(19, Languages.Chinese_Traditional);
            AddMappping(20, Languages.Urdu);
            AddMappping(21, Languages.Hindi);
            AddMappping(22, Languages.Thai);
            AddMappping(23, Languages.Korean);
            AddMappping(24, Languages.Lithuanian);
            AddMappping(25, Languages.Polish);
            AddMappping(26, Languages.Hungarian);
            AddMappping(27, Languages.Estonian);
            AddMappping(28, Languages.Latvian);
            AddMappping(29, Languages.Sami);
            AddMappping(30, Languages.Faroese);
            AddMappping(31, Languages.Farsi);
            AddMappping(32, Languages.Russian);
            AddMappping(33, Languages.Chinese_Simplified);
            AddMappping(34, Languages.Flemish);
            AddMappping(35, Languages.Irish);
            AddMappping(36, Languages.Albanian);
            AddMappping(37, Languages.Romanian);
            AddMappping(38, Languages.Czech);
            AddMappping(39, Languages.Slovak);
            AddMappping(40, Languages.Slovenian);
            AddMappping(41, Languages.Yiddish);
            AddMappping(42, Languages.Serbian);
            AddMappping(43, Languages.Macedonian);
            AddMappping(44, Languages.Bulgarian);
            AddMappping(45, Languages.Ukrainian);
            AddMappping(46, Languages.Byelorussian);
            AddMappping(47, Languages.Uzbek);
            AddMappping(48, Languages.Kazakh);
            AddMappping(49, Languages.Azeri_Cyrillic);
            AddMappping(50, Languages.Azeri_Arabic);
            AddMappping(51, Languages.Armenian);
            AddMappping(52, Languages.Georgian);
            AddMappping(53, Languages.Moldavian);
            AddMappping(54, Languages.Kirghiz);
            AddMappping(55, Languages.Tajiki);
            AddMappping(56, Languages.Turkmen);
            AddMappping(57, Languages.Mongolian_Traditional);
            AddMappping(58, Languages.Mongolian_Cyrillic);
            AddMappping(59, Languages.Pashto);
            AddMappping(60, Languages.Kurdish);
            AddMappping(61, Languages.Kashmiri);
            AddMappping(62, Languages.Sindhi);
            AddMappping(63, Languages.Tibetan);
            AddMappping(64, Languages.Nepali);
            AddMappping(65, Languages.Sanskrit);
            AddMappping(66, Languages.Marathi);
            AddMappping(67, Languages.Bengali);
            AddMappping(68, Languages.Assamese);
            AddMappping(69, Languages.Gujarati);
            AddMappping(70, Languages.Punjabi);
            AddMappping(71, Languages.Oriya);
            AddMappping(72, Languages.Malayalam);
            AddMappping(73, Languages.Kannada);
            AddMappping(74, Languages.Tamil);
            AddMappping(75, Languages.Telugu);
            AddMappping(76, Languages.Sinhalese);
            AddMappping(77, Languages.Burmese);
            AddMappping(78, Languages.Khmer);
            AddMappping(79, Languages.Lao);
            AddMappping(80, Languages.Vietnamese);
            AddMappping(81, Languages.Indonesian);
            AddMappping(82, Languages.Tagalog);
            AddMappping(83, Languages.Malay_Roman);
            AddMappping(84, Languages.Malay_Arabic);
            AddMappping(85, Languages.Amharic);
            AddMappping(86, Languages.Tigrinya);
            AddMappping(87, Languages.Galla);
            AddMappping(88, Languages.Somali);
            AddMappping(89, Languages.Swahili);
            AddMappping(90, Languages.Kinyarwanda);
            AddMappping(91, Languages.Rundi);
            AddMappping(128, Languages.Welsh);
            AddMappping(129, Languages.Basque);
            AddMappping(130, Languages.Catalan);
            AddMappping(131, Languages.Latin);
            AddMappping(132, Languages.Quechua);
            AddMappping(133, Languages.Gurani);
            AddMappping(134, Languages.Aymara);
            AddMappping(135, Languages.Tatar);
            AddMappping(136, Languages.Uighur);
            AddMappping(137, Languages.Dzongkha);
            AddMappping(138, Languages.Javanese_Roman);
            AddMappping(139, Languages.Sundanese_Roman);
            AddMappping(140, Languages.Galician);
            AddMappping(141, Languages.Afrikaans);
            AddMappping(142, Languages.Breton);
            AddMappping(143, Languages.Inuktitut);
            AddMappping(144, Languages.Galician);
            AddMappping(145, Languages.Galician);
            AddMappping(146, Languages.Irish);
            AddMappping(147, Languages.Tongan);
            AddMappping(148, Languages.Greek_Polytonic);
            AddMappping(149, Languages.Greenlandic);
            AddMappping(150, Languages.Azeri_Latin);
        }

        public static IDictionary<int, LanguageMapping> Mappings
        {
            get
            {
                if (_mappings.Count() == 0)
                {
                    CreateMappings();
                }
                return _mappings;
            }
        }
    }
}
