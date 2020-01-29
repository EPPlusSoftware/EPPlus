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
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// The calendar to be used.
    /// </summary>
    public enum eCalendarType
    {
        /// <summary>
        /// The Gregorian calendar
        /// </summary>
        Gregorian,
        /// <summary>
        /// The Gregorian calendar, as defined in ISO 8601. Arabic. This calendar should be localized into the appropriate language.        
        /// /// </summary>
        GregorianArabic,
        /// <summary>
        /// The Gregorian calendar, as defined in ISO 8601. Middle East French.
        /// </summary>
        GregorianMeFrench,
        /// <summary>
        ///  The Gregorian calendar, as defined in ISO 8601. English.
        /// </summary>
        GregorianUs,
        /// <summary>
        /// The Gregorian calendar, as defined in ISO 8601. English strings in the corresponding Arabic characters. The Arabic transliteration of the English for the Gregoriancalendar.
        /// </summary>
        GregorianXlitEnglish,
        /// <summary>
        /// The Gregorian calendar, as defined in ISO 8601. French strings in the corresponding Arabic characters. The Arabic transliteration of the French for the Gregoriancalendar.
        /// </summary>
        GregorianXlitFrench,
        /// <summary>
        /// The Hijri lunar calendar, as described by the Kingdom of Saudi Arabia, Ministry of Islamic Affairs, Endowments, Da‘wah and Guidance
        /// </summary>
        Hijri,
        /// <summary>
        /// The Hebrew lunar calendar, as described by the Gauss formula for Passover [Har'El, Zvi] and The Complete Restatement of Oral Law(Mishneh Torah).
        /// </summary>
        Hebrew,
        /// <summary>
        /// The Japanese Emperor Era calendar, as described by Japanese Industrial Standard JIS X 0301.
        /// </summary>
        Japan,
        /// <summary>
        /// The Korean Tangun Era calendar, as described by Korean Law Enactment No. 4
        /// </summary>
        Korea,
        /// <summary>
        /// No calendar
        /// </summary>
        None,
        /// <summary>
        /// The Saka Era calendar, as described by the Calendar Reform Committee of India, as part of the Indian Ephemeris and Nautical Almanac
        /// </summary>
        Taiwan,
        /// <summary>
        /// The Thai calendar, as defined by the Royal Decree of H.M. King Vajiravudh (Rama VI) in Royal Gazette B. E. 2456 (1913 A.D.) and by the decree of Prime Minister Phibunsongkhram (1941 A.D.) to start the year on the Gregorian January 1 and to map year zero to Gregorian year 543 B.C.
        /// </summary>
        Thai
    }
}