using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.FontLocalization;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.FontLocalization
{
    internal static class WindowsLanguageMappings
    {
        private static IDictionary<int, LanguageMapping> _mappings = new Dictionary<int, LanguageMapping>();
        private static void AddMapping(int hexNumber, Languages language)
        {
            lock (_mappings)
            {
                if (_mappings.ContainsKey(hexNumber)) return;

                var mapping = LanguageMapping.Create(hexNumber, language);
                _mappings.Add(mapping.code, mapping);
            }
        }
        private static void CreateMappings()
        {

            AddMapping(0x0436, Languages.Afrikaans);
            AddMapping(0x041C, Languages.Albanian);
            AddMapping(0x0484, Languages.Alsatian);
            AddMapping(0x045E, Languages.Amharic);
            
            AddMapping(0x1401, Languages.Arabic);
            AddMapping(0x3C01, Languages.Arabic);
            AddMapping(0x0C01, Languages.Arabic);
            AddMapping(0x0801, Languages.Arabic);
            AddMapping(0x2C01, Languages.Arabic);
            AddMapping(0x3401, Languages.Arabic);
            AddMapping(0x3001, Languages.Arabic);
            AddMapping(0x1001, Languages.Arabic);
            AddMapping(0x1801, Languages.Arabic);
            AddMapping(0x2001, Languages.Arabic);
            AddMapping(0x4001, Languages.Arabic);
            AddMapping(0x0401, Languages.Arabic);
            AddMapping(0x2801, Languages.Arabic);
            AddMapping(0x1C01, Languages.Arabic);
            AddMapping(0x3801, Languages.Arabic);
            AddMapping(0x2401, Languages.Arabic);
            
            AddMapping(0x042B, Languages.Armenian);
            AddMapping(0x044D, Languages.Assamese);
            AddMapping(0x082C, Languages.Azeri_Cyrillic);
            AddMapping(0x042C, Languages.Azeri_Latin);
            AddMapping(0x046D, Languages.Bashkir);
            AddMapping(0x042D, Languages.Basque);
            AddMapping(0x0423, Languages.Belarusian);
            AddMapping(0x0845, Languages.Bengali);
            AddMapping(0x0445, Languages.Bengali);
            AddMapping(0x201A, Languages.Bosnian_Cyrillic);
            AddMapping(0x141A, Languages.Bosnian_Latin);

            AddMapping(0x047E, Languages.Breton);
            AddMapping(0x0402, Languages.Bulgarian);
            AddMapping(0x0403, Languages.Catalan);
            AddMapping(0x1404, Languages.Chinese);
            AddMapping(0x0804, Languages.Chinese);
            AddMapping(0x1004, Languages.Chinese);
            AddMapping(0x0404, Languages.Chinese);
            AddMapping(0x0483, Languages.Corsican);
            AddMapping(0x041A, Languages.Croatian);

            AddMapping(0x101A, Languages.Croatian_Latin);
            AddMapping(0x0405, Languages.Czech );
            AddMapping(0x0406, Languages.Danish);
            AddMapping(0x048C, Languages.Dari);
            AddMapping(0x0465, Languages.Divehi);
            AddMapping(0x0813, Languages.Dutch);
            AddMapping(0x0413, Languages.Dutch);

            AddMapping(0x0C09, Languages.English);
            AddMapping(0x2809, Languages.English);
            AddMapping(0x1009, Languages.English);
            AddMapping(0x2409, Languages.English);
            AddMapping(0x4009, Languages.English);
            AddMapping(0x1809, Languages.English);
            AddMapping(0x2009, Languages.English);
            AddMapping(0x4409, Languages.English);
            AddMapping(0x1409, Languages.English);
            AddMapping(0x3409, Languages.English);
            AddMapping(0x4809, Languages.English);
            AddMapping(0x1C09, Languages.English);
            AddMapping(0x2C09, Languages.English);
            AddMapping(0x0809, Languages.English);
            AddMapping(0x0409, Languages.English);
            AddMapping(0x3009, Languages.English);

            AddMapping(0x0425, Languages.Estonian);
            AddMapping(0x0438, Languages.Faroese);
            AddMapping(0x0464, Languages.Filipino);
            AddMapping(0x040B, Languages.Finnish);
            AddMapping(0x080C, Languages.French);
            AddMapping(0x0C0C, Languages.French);
            AddMapping(0x040C, Languages.French);
            AddMapping(0x140c, Languages.French);
            AddMapping(0x180C, Languages.French);
            AddMapping(0x100C, Languages.French);
            AddMapping(0x0462, Languages.Frisian);

            AddMapping(0x0456, Languages.Galician);
            AddMapping(0x0437, Languages.Georgian);
            AddMapping(0x0C07, Languages.German);
            AddMapping(0x0407, Languages.German);
            AddMapping(0x1407, Languages.German);
            AddMapping(0x1007, Languages.German);
            AddMapping(0x0807, Languages.German);
            AddMapping(0x0408, Languages.Greek);
            AddMapping(0x046F, Languages.Greenlandic);
            AddMapping(0x0447, Languages.Gujarati);

            AddMapping(0x0468, Languages.Hausa_Latin);
            AddMapping(0x040D, Languages.Hebrew);
            AddMapping(0x0439, Languages.Hindi);
            AddMapping(0x040E, Languages.Hungarian);
            AddMapping(0x040F, Languages.Icelandic);
            AddMapping(0x0470, Languages.Igbo);
            AddMapping(0x0421, Languages.Indonesian);
            AddMapping(0x045D, Languages.Inuktitut);
            AddMapping(0x085D, Languages.Inuktitut_Latin);
            AddMapping(0x083C, Languages.Irish);
            AddMapping(0x0434, Languages.isiXhosa);
            AddMapping(0x0435, Languages.isiZulu);

            AddMapping(0x0410, Languages.Italian);
            AddMapping(0x0810, Languages.Italian);
            AddMapping(0x0411, Languages.Japanese);
            AddMapping(0x044B, Languages.Kannada);
            AddMapping(0x043F, Languages.Kazakh);
            AddMapping(0x0453, Languages.Khmer);
            AddMapping(0x0486, Languages.Kiche);
            AddMapping(0x0487, Languages.Kinyarwanda);
            AddMapping(0x0441, Languages.Kiswahili);
            AddMapping(0x0457, Languages.Konkani);

            AddMapping(0x0412, Languages.Korean);
            AddMapping(0x0440, Languages.Kyrgyz);
            AddMapping(0x0454, Languages.Lao);
            AddMapping(0x0426, Languages.Latvian);
            AddMapping(0x0427, Languages.Lithuanian);
            AddMapping(0x082E, Languages.LowerSorbian);
            AddMapping(0x046E, Languages.Luxembourgish);
            AddMapping(0x042F, Languages.Macedonian);
            AddMapping(0x083E, Languages.Malay);
            AddMapping(0x043E, Languages.Malay);

            AddMapping(0x044C, Languages.Malayalam);
            AddMapping(0x043A, Languages.Maltese);
            AddMapping(0x0481, Languages.Maori);
            AddMapping(0x047A, Languages.Mapudungun);
            AddMapping(0x044E, Languages.Marathi);
            AddMapping(0x047C, Languages.Mohawk);
            AddMapping(0x0450, Languages.Mongolian_Cyrillic);
            AddMapping(0x0850, Languages.Mongolian_Traditional);
            AddMapping(0x0461, Languages.Nepali);

            AddMapping(0x0414, Languages.Norwegian_Bokmal);
            AddMapping(0x0814, Languages.Norwegian_Nynorsk);
            AddMapping(0x0482, Languages.Occitan);
            AddMapping(0x0448, Languages.Odia_formerlyOriya);
            AddMapping(0x0463, Languages.Pashto);
            AddMapping(0x0415, Languages.Polish);
            AddMapping(0x0416, Languages.Portuguese);
            AddMapping(0x0816, Languages.Portuguese);

            AddMapping(0x0446, Languages.Punjabi);
            AddMapping(0x046B, Languages.Quechua);
            AddMapping(0x086B, Languages.Quechua);
            AddMapping(0x0C6B, Languages.Quechua);
            AddMapping(0x0418, Languages.Romanian);
            AddMapping(0x0417, Languages.Romansh);
            AddMapping(0x0419, Languages.Russian);
            AddMapping(0x243B, Languages.Sami_Inari);
            AddMapping(0x103B, Languages.Sami_Lule);
            AddMapping(0x143B, Languages.Sami_Lule);

            AddMapping(0x0C3B, Languages.Sami_Northern);
            AddMapping(0x043B, Languages.Sami_Northern);
            AddMapping(0x083B, Languages.Sami_Northern);
            AddMapping(0x203B, Languages.Sami_Skolt);
            AddMapping(0x183B, Languages.Sami_Southern);
            AddMapping(0x1C3B, Languages.Sami_Southern);
            AddMapping(0x044F, Languages.Sanskrit);
            AddMapping(0x1C1A, Languages.Serbian_Cyrillic);
            AddMapping(0x0C1A, Languages.Serbian_Cyrillic);
            AddMapping(0x181A, Languages.Serbian_Latin);
            AddMapping(0x081A, Languages.Serbian_Latin);

            AddMapping(0x046C, Languages.Sesotho_saLeboa);
            AddMapping(0x0432, Languages.Setswana);
            AddMapping(0x045B, Languages.Sinhala);
            AddMapping(0x041B, Languages.Slovak);
            AddMapping(0x0424, Languages.Slovenian);
            AddMapping(0x2C0A, Languages.Spanish);
            AddMapping(0x400A, Languages.Spanish);
            AddMapping(0x340A, Languages.Spanish);
            AddMapping(0x240A, Languages.Spanish);
            AddMapping(0x140A, Languages.Spanish);
            AddMapping(0x1C0A, Languages.Spanish);
            AddMapping(0x300A, Languages.Spanish);
            AddMapping(0x440A, Languages.Spanish);
            AddMapping(0x100A, Languages.Spanish);
            AddMapping(0x480A, Languages.Spanish);
            AddMapping(0x080A, Languages.Spanish);
            AddMapping(0x4C0A, Languages.Spanish);
            AddMapping(0x180A, Languages.Spanish);
            AddMapping(0x3C0A, Languages.Spanish);
            AddMapping(0x280A, Languages.Spanish);
            AddMapping(0x500A, Languages.Spanish);
            AddMapping(0x0C0A, Languages.Spanish_ModernSort);
            AddMapping(0x040A, Languages.Spanish_TraditionalSort);
            AddMapping(0x540A, Languages.Spanish);
            AddMapping(0x380A, Languages.Spanish);
            AddMapping(0x200A, Languages.Spanish);

            AddMapping(0x081D, Languages.Swedish);
            AddMapping(0x041D, Languages.Swedish);
            AddMapping(0x045A, Languages.Syriac);
            AddMapping(0x0428, Languages.Tajik_Cyrillic);
            AddMapping(0x085F, Languages.Tamazight_Latin);
            AddMapping(0x0449, Languages.Tamil);
            AddMapping(0x0444, Languages.Tatar);
            AddMapping(0x044A, Languages.Telugu);
            AddMapping(0x041E, Languages.Thai);
            AddMapping(0x0451, Languages.Tibetan);
            AddMapping(0x041F, Languages.Turkish);
            AddMapping(0x0442, Languages.Turkmen);
            AddMapping(0x0480, Languages.Uighur);
            AddMapping(0x0422, Languages.Ukrainian);

            AddMapping(0x042E, Languages.UpperSorbian);
            AddMapping(0x0420, Languages.Urdu);
            AddMapping(0x0843, Languages.Uzbek_Cyrillic);
            AddMapping(0x0443, Languages.Uzbek_Latin);
            AddMapping(0x042A, Languages.Vietnamese);
            AddMapping(0x0452, Languages.Welsh);
            AddMapping(0x0488, Languages.Wolof);
            AddMapping(0x0485, Languages.Yakut);
            AddMapping(0x0478, Languages.Yi);
            AddMapping(0x046A, Languages.Yoruba);
        }

        public static IDictionary<int, LanguageMapping> Mappings
        {
            get
            {
                if(_mappings.Count() == 0)
                {
                    CreateMappings();
                }
                return _mappings;
            }
        }
    }
}
