namespace FontLab1.FontLocalization
{
    internal class LanguageMapping
    {
        public int code { get; set; }

        public Languages Language { get; set; }

        public static LanguageMapping Create(int code, Languages language)
        {
            return new LanguageMapping
            {
                code = code,
                Language = language
            };
        }

        public override string ToString()
        {
            return Language.ToString();
        }
    }
}
