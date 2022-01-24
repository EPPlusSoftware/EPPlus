using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.FontLocalization
{
    public static class WindowsLanguageMappings
    {
        private static IDictionary<int, LanguageMapping> _mappings = new Dictionary<int, LanguageMapping>();
        private static void AddMappping(int hexNumber, Languages language)
        {
            var mapping = LanguageMapping.Create(hexNumber, language);
            _mappings.Add(mapping.code, mapping);
        }
        private static void CreateMappings()
        {

            AddMappping(0x0436, Languages.Afrikaans);
            AddMappping(0x041C, Languages.Albanian);
            AddMappping(0x0484, Languages.Alsatian);
            AddMappping(0x045E, Languages.Amharic);
            
            AddMappping(0x1401, Languages.Arabic);
            AddMappping(0x3C01, Languages.Arabic);
            AddMappping(0x0C01, Languages.Arabic);
            AddMappping(0x0801, Languages.Arabic);
            AddMappping(0x2C01, Languages.Arabic);
            AddMappping(0x3401, Languages.Arabic);
            AddMappping(0x3001, Languages.Arabic);
            AddMappping(0x1001, Languages.Arabic);
            AddMappping(0x1801, Languages.Arabic);
            AddMappping(0x2001, Languages.Arabic);
            AddMappping(0x4001, Languages.Arabic);
            AddMappping(0x0401, Languages.Arabic);
            AddMappping(0x2801, Languages.Arabic);
            AddMappping(0x1C01, Languages.Arabic);
            AddMappping(0x3801, Languages.Arabic);
            AddMappping(0x2401, Languages.Arabic);
            
            AddMappping(0x042B, Languages.Armenian);
            AddMappping(0x044D, Languages.Assamese);
            AddMappping(0x082C, Languages.Azeri_Cyrillic);
            AddMappping(0x042C, Languages.Azeri_Latin);
            AddMappping(0x046D, Languages.Bashkir);
            AddMappping(0x042D, Languages.Basque);
            AddMappping(0x0423, Languages.Belarusian);
            AddMappping(0x0845, Languages.Bengali);
            AddMappping(0x0445, Languages.Bengali);
            AddMappping(0x201A, Languages.Bosnian_Cyrillic);
            AddMappping(0x141A, Languages.Bosnian_Latin);

            AddMappping(0x047E, Languages.Breton);
            AddMappping(0x0402, Languages.Bulgarian);
            AddMappping(0x0403, Languages.Catalan);
            AddMappping(0x1404, Languages.Chinese);
            AddMappping(0x0804, Languages.Chinese);
            AddMappping(0x1004, Languages.Chinese);
            AddMappping(0x0404, Languages.Chinese);
            AddMappping(0x0483, Languages.Corsican);
            AddMappping(0x041A, Languages.Croatian);

            AddMappping(0x101A, Languages.Croatian_Latin);
            AddMappping(0x0405, Languages.Czech );
            AddMappping(0x0406, Languages.Danish);
            AddMappping(0x048C, Languages.Dari);
            AddMappping(0x0465, Languages.Divehi);
            AddMappping(0x0813, Languages.Dutch);
            AddMappping(0x0413, Languages.Dutch);

            AddMappping(0x0C09, Languages.English);
            AddMappping(0x2809, Languages.English);
            AddMappping(0x1009, Languages.English);
            AddMappping(0x2409, Languages.English);
            AddMappping(0x4009, Languages.English);
            AddMappping(0x1809, Languages.English);
            AddMappping(0x2009, Languages.English);
            AddMappping(0x4409, Languages.English);
            AddMappping(0x1409, Languages.English);
            AddMappping(0x3409, Languages.English);
            AddMappping(0x4809, Languages.English);
            AddMappping(0x1C09, Languages.English);
            AddMappping(0x2C09, Languages.English);
            AddMappping(0x0809, Languages.English);
            AddMappping(0x0409, Languages.English);
            AddMappping(0x3009, Languages.English);

            AddMappping(0x0425, Languages.Estonian);
            AddMappping(0x0438, Languages.Faroese);
            AddMappping(0x0464, Languages.Filipino);
            AddMappping(0x040B, Languages.Finnish);
            AddMappping(0x080C, Languages.French);
            AddMappping(0x0C0C, Languages.French);
            AddMappping(0x040C, Languages.French);
            AddMappping(0x140c, Languages.French);
            AddMappping(0x180C, Languages.French);
            AddMappping(0x100C, Languages.French);
            AddMappping(0x0462, Languages.Frisian);

            AddMappping(0x0456, Languages.Galician);
            AddMappping(0x0437, Languages.Georgian);
            AddMappping(0x0C07, Languages.German);
            AddMappping(0x0407, Languages.German);
            AddMappping(0x1407, Languages.German);
            AddMappping(0x1007, Languages.German);
            AddMappping(0x0807, Languages.German);
            AddMappping(0x0408, Languages.Greek);
            AddMappping(0x046F, Languages.Greenlandic);
            AddMappping(0x0447, Languages.Gujarati);

            AddMappping(0x0468, Languages.Hausa_Latin);
            AddMappping(0x040D, Languages.Hebrew);
            AddMappping(0x0439, Languages.Hindi);
            AddMappping(0x040E, Languages.Hungarian);
            AddMappping(0x040F, Languages.Icelandic);
            AddMappping(0x0470, Languages.Igbo);
            AddMappping(0x0421, Languages.Indonesian);
            AddMappping(0x045D, Languages.Inuktitut);
            AddMappping(0x085D, Languages.Inuktitut_Latin);
            AddMappping(0x083C, Languages.Irish);
            AddMappping(0x0434, Languages.isiXhosa);
            AddMappping(0x0435, Languages.isiZulu);

            AddMappping(0x0410, Languages.Italian);
            AddMappping(0x0810, Languages.Italian);
            AddMappping(0x0411, Languages.Japanese);
            AddMappping(0x044B, Languages.Kannada);
            AddMappping(0x043F, Languages.Kazakh);
            AddMappping(0x0453, Languages.Khmer);
            AddMappping(0x0486, Languages.Kiche);
            AddMappping(0x0487, Languages.Kinyarwanda);
            AddMappping(0x0441, Languages.Kiswahili);
            AddMappping(0x0457, Languages.Konkani);

            AddMappping(0x0412, Languages.Korean);
            AddMappping(0x0440, Languages.Kyrgyz);
            AddMappping(0x0454, Languages.Lao);
            AddMappping(0x0426, Languages.Latvian);
            AddMappping(0x0427, Languages.Lithuanian);
            AddMappping(0x082E, Languages.LowerSorbian);
            AddMappping(0x046E, Languages.Luxembourgish);
            AddMappping(0x042F, Languages.Macedonian);
            AddMappping(0x083E, Languages.Malay);
            AddMappping(0x043E, Languages.Malay);

            AddMappping(0x044C, Languages.Malayalam);
            AddMappping(0x043A, Languages.Maltese);
            AddMappping(0x0481, Languages.Maori);
            AddMappping(0x047A, Languages.Mapudungun);
            AddMappping(0x044E, Languages.Marathi);
            AddMappping(0x047C, Languages.Mohawk);
            AddMappping(0x0450, Languages.Mongolian_Cyrillic);
            AddMappping(0x0850, Languages.Mongolian_Traditional);
            AddMappping(0x0461, Languages.Nepali);

            AddMappping(0x0414, Languages.Norwegian_Bokmal);
            AddMappping(0x0814, Languages.Norwegian_Nynorsk);
            AddMappping(0x0482, Languages.Occitan);
            AddMappping(0x0448, Languages.Odia_formerlyOriya);
            AddMappping(0x0463, Languages.Pashto);
            AddMappping(0x0415, Languages.Polish);
            AddMappping(0x0416, Languages.Portuguese);
            AddMappping(0x0816, Languages.Portuguese);

            AddMappping(0x0446, Languages.Punjabi);
            AddMappping(0x046B, Languages.Quechua);
            AddMappping(0x086B, Languages.Quechua);
            AddMappping(0x0C6B, Languages.Quechua);
            AddMappping(0x0418, Languages.Romanian);
            AddMappping(0x0417, Languages.Romansh);
            AddMappping(0x0419, Languages.Russian);
            AddMappping(0x243B, Languages.Sami_Inari);
            AddMappping(0x103B, Languages.Sami_Lule);
            AddMappping(0x143B, Languages.Sami_Lule);

            AddMappping(0x0C3B, Languages.Sami_Northern);
            AddMappping(0x043B, Languages.Sami_Northern);
            AddMappping(0x083B, Languages.Sami_Northern);
            AddMappping(0x203B, Languages.Sami_Skolt);
            AddMappping(0x183B, Languages.Sami_Southern);
            AddMappping(0x1C3B, Languages.Sami_Southern);
            AddMappping(0x044F, Languages.Sanskrit);
            AddMappping(0x1C1A, Languages.Serbian_Cyrillic);
            AddMappping(0x0C1A, Languages.Serbian_Cyrillic);
            AddMappping(0x181A, Languages.Serbian_Latin);
            AddMappping(0x081A, Languages.Serbian_Latin);

            AddMappping(0x046C, Languages.Sesotho_saLeboa);
            AddMappping(0x0432, Languages.Setswana);
            AddMappping(0x045B, Languages.Sinhala);
            AddMappping(0x041B, Languages.Slovak);
            AddMappping(0x0424, Languages.Slovenian);
            AddMappping(0x2C0A, Languages.Spanish);
            AddMappping(0x400A, Languages.Spanish);
            AddMappping(0x340A, Languages.Spanish);
            AddMappping(0x240A, Languages.Spanish);
            AddMappping(0x140A, Languages.Spanish);
            AddMappping(0x1C0A, Languages.Spanish);
            AddMappping(0x300A, Languages.Spanish);
            AddMappping(0x440A, Languages.Spanish);
            AddMappping(0x100A, Languages.Spanish);
            AddMappping(0x480A, Languages.Spanish);
            AddMappping(0x080A, Languages.Spanish);
            AddMappping(0x4C0A, Languages.Spanish);
            AddMappping(0x180A, Languages.Spanish);
            AddMappping(0x3C0A, Languages.Spanish);
            AddMappping(0x280A, Languages.Spanish);
            AddMappping(0x500A, Languages.Spanish);
            AddMappping(0x0C0A, Languages.Spanish_ModernSort);
            AddMappping(0x040A, Languages.Spanish_TraditionalSort);
            AddMappping(0x540A, Languages.Spanish);
            AddMappping(0x380A, Languages.Spanish);
            AddMappping(0x200A, Languages.Spanish);

            AddMappping(0x081D, Languages.Swedish);
            AddMappping(0x041D, Languages.Swedish);
            AddMappping(0x045A, Languages.Syriac);
            AddMappping(0x0428, Languages.Tajik_Cyrillic);
            AddMappping(0x085F, Languages.Tamazight_Latin);
            AddMappping(0x0449, Languages.Tamil);
            AddMappping(0x0444, Languages.Tatar);
            AddMappping(0x044A, Languages.Telugu);
            AddMappping(0x041E, Languages.Thai);
            AddMappping(0x0451, Languages.Tibetan);
            AddMappping(0x041F, Languages.Turkish);
            AddMappping(0x0442, Languages.Turkmen);
            AddMappping(0x0480, Languages.Uighur);
            AddMappping(0x0422, Languages.Ukrainian);

            AddMappping(0x042E, Languages.UpperSorbian);
            AddMappping(0x0420, Languages.Urdu);
            AddMappping(0x0843, Languages.Uzbek_Cyrillic);
            AddMappping(0x0443, Languages.Uzbek_Latin);
            AddMappping(0x042A, Languages.Vietnamese);
            AddMappping(0x0452, Languages.Welsh);
            AddMappping(0x0488, Languages.Wolof);
            AddMappping(0x0485, Languages.Yakut);
            AddMappping(0x0478, Languages.Yi);
            AddMappping(0x046A, Languages.Yoruba);
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
