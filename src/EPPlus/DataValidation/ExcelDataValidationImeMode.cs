namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Mode for east-asian languages who use Input Method Editors(IME)
    /// </summary>
    public enum ExcelDataValidationImeMode
    {
        /// <summary>
        /// Default. Has no effect on IME
        /// </summary>
        NoControl = 0,
        /// <summary>
        /// Forces IME mode to OFF
        /// </summary>
        Off = 1,
        /// <summary>
        /// Forces the IMEmode to be on when first selecting the cell
        /// </summary>
        On = 2,
        /// <summary>
        /// IME mode is disabled when cell is selected
        /// </summary>
        Disabled = 3,
        /// <summary>
        /// Forces on and Hiragana (only applies if Japanese IME)
        /// </summary>
        Hiragana = 4,
        /// <summary>
        /// Forces on and full-width katakana
        /// </summary>
        FullKatakana = 5,
        /// <summary>
        /// Forces on and half-width katakana
        /// </summary>
        HalfKatakana = 6,
        /// <summary>
        /// Forces on and Alpha-Numeric IME
        /// </summary>
        FullAlpha = 7,
        /// <summary>
        /// Forces on and half-width alpha-numeric
        /// </summary>
        HalfAlpha = 8,
        /// <summary>
        /// Forces on and Full-width Hangul if Korean IME
        /// </summary>
        FullHangul = 9,
        /// <summary>
        /// Forces on and half-width Hangul
        /// </summary>
        HalfHangul = 10
    }
}
