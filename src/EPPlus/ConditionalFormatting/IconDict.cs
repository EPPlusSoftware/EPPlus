using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class IconDict
    {
        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeArrows =
    new List<eExcelconditionalFormattingCustomIcon>()
    { eExcelconditionalFormattingCustomIcon.RedDownArrow, eExcelconditionalFormattingCustomIcon.YellowSideArrow, eExcelconditionalFormattingCustomIcon.GreenUpArrow};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeArrowsGray =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.GrayDownArrow, eExcelconditionalFormattingCustomIcon.YellowSideArrow, eExcelconditionalFormattingCustomIcon.GrayUpArrow};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeFlags =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedFlag, eExcelconditionalFormattingCustomIcon.YellowFlag, eExcelconditionalFormattingCustomIcon.GreenFlag};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeSigns =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedDiamond, eExcelconditionalFormattingCustomIcon.YellowTriangle, eExcelconditionalFormattingCustomIcon.GreenCircle};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeSymbols =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedCrossSymbol, eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol, eExcelconditionalFormattingCustomIcon.GreenCheckSymbol};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeSymbols2 =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedCross, eExcelconditionalFormattingCustomIcon.YellowExclamation, eExcelconditionalFormattingCustomIcon.GreenCheck};

        static readonly List<eExcelconditionalFormattingCustomIcon> TrafficLights1 =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedCircle, eExcelconditionalFormattingCustomIcon.YellowCircle, eExcelconditionalFormattingCustomIcon.GreenCircle};

        static readonly List<eExcelconditionalFormattingCustomIcon> TrafficLights2 =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedTrafficLight, eExcelconditionalFormattingCustomIcon.YellowTrafficLight, eExcelconditionalFormattingCustomIcon.GreenTrafficLight};

        static readonly List<eExcelconditionalFormattingCustomIcon> Stars =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.SilverStar, eExcelconditionalFormattingCustomIcon.HalfGoldStar, eExcelconditionalFormattingCustomIcon.GoldStar};

        static readonly List<eExcelconditionalFormattingCustomIcon> ThreeTriangles =
            new List<eExcelconditionalFormattingCustomIcon>()
            { eExcelconditionalFormattingCustomIcon.RedDownTriangle, eExcelconditionalFormattingCustomIcon.YellowDash, eExcelconditionalFormattingCustomIcon.GreenUpTriangle};
        //4 iconsets:

        static readonly List<eExcelconditionalFormattingCustomIcon> FourArrows =
            new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.RedDownArrow, eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow, eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow ,eExcelconditionalFormattingCustomIcon.GreenUpArrow };

        static readonly List<eExcelconditionalFormattingCustomIcon> FourArrowsGray =
            new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.GrayDownArrow, eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow, eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow ,eExcelconditionalFormattingCustomIcon.GrayUpArrow };

        static readonly List<eExcelconditionalFormattingCustomIcon> FourRating =
            new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar, eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars, eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars ,eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars };

        static readonly List<eExcelconditionalFormattingCustomIcon> FourRedToBlack =
            new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.BlackCircle, eExcelconditionalFormattingCustomIcon.GrayCircle, eExcelconditionalFormattingCustomIcon.PinkCircle ,eExcelconditionalFormattingCustomIcon.RedCircle };

        static readonly List<eExcelconditionalFormattingCustomIcon> FourTrafficLights =
            new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder, eExcelconditionalFormattingCustomIcon.RedCircle, eExcelconditionalFormattingCustomIcon.YellowCircle ,eExcelconditionalFormattingCustomIcon.GreenCircle };

        //5 iconsets:
        static readonly List<eExcelconditionalFormattingCustomIcon> FiveArrows =
             new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.RedDownArrow, eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow, eExcelconditionalFormattingCustomIcon.YellowSideArrow, eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow ,eExcelconditionalFormattingCustomIcon.GreenUpArrow };

        static readonly List<eExcelconditionalFormattingCustomIcon> FiveArrowsGray =
             new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.GrayDownArrow, eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow, eExcelconditionalFormattingCustomIcon.GraySideArrow, eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow ,eExcelconditionalFormattingCustomIcon.GrayUpArrow };

        static readonly List<eExcelconditionalFormattingCustomIcon> FiveQuarters =
             new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.WhiteCircle, eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters, eExcelconditionalFormattingCustomIcon.CircleWithTwoWhiteQuarters, eExcelconditionalFormattingCustomIcon.CircleWithOneWhiteQuarter ,eExcelconditionalFormattingCustomIcon.BlackCircle };

        static readonly List<eExcelconditionalFormattingCustomIcon> FiveRating =
             new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars, eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar, eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars, eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars ,eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars };

        static readonly List<eExcelconditionalFormattingCustomIcon> FiveBoxes =
             new List<eExcelconditionalFormattingCustomIcon>()
            {eExcelconditionalFormattingCustomIcon.ZeroFilledBoxes, eExcelconditionalFormattingCustomIcon.OneFilledBox, eExcelconditionalFormattingCustomIcon.TwoFilledBoxes, eExcelconditionalFormattingCustomIcon.ThreeFilledBoxes ,eExcelconditionalFormattingCustomIcon.FourFilledBoxes };


        internal static readonly Dictionary<string, List<eExcelconditionalFormattingCustomIcon>> IconSets
            = new Dictionary<string, List<eExcelconditionalFormattingCustomIcon>>
            {
                {"3Arrows", ThreeArrows },
                {"3ArrowsGray", ThreeArrowsGray },
                {"3Flags", ThreeFlags },
                {"3Signs", ThreeSigns },
                {"3Symbols", ThreeSymbols },
                {"3Symbols2", ThreeSymbols2 },
                {"3TrafficLights1", TrafficLights1 },
                {"3TrafficLights2", TrafficLights2 },
                {"3Stars", Stars },
                {"3Triangles", ThreeTriangles },
                {"4Arrows", FourArrows },
                {"4ArrowsGray", FourArrowsGray },
                {"4Rating", FourRating },
                {"4RedToBlack", FourRedToBlack },
                {"4TrafficLights", FourTrafficLights },
                {"5Arrows", FiveArrows },
                {"5ArrowsGray", FiveArrowsGray },
                {"5Quarters", FiveQuarters },
                {"5Rating", FiveRating },
                {"5Boxes", FiveBoxes },

            };

        //internal static readonly Dictionary<string, eExcelconditionalFormattingCustomIcon[]> IconSets =
        //    new Dictionary<string, eExcelconditionalFormattingCustomIcon[]>
        //    {
        //        {"3Arrows", GetIconSet("3Arrows") },
        //        {"3ArrowsGray", GetIconSet("3ArrowsGray") },
        //        {"3Flags", GetIconSet("3Flags") },
        //        {"3TrafficLights1", GetIconSet("3TrafficLights1") },
        //        {"3TrafficLights2", GetIconSet("3TrafficLights2") },
        //        {"3Signs", GetIconSet("3Signs") },
        //        {"3Symbols", GetIconSet("3Symbols") },
        //        {"3Symbols2", GetIconSet("3Symbols2") },
        //        {"3Stars", GetIconSet("3Stars") },
        //        {"3Triangles", GetIconSet("3Triangles") },
        //        {"4Arrows", GetIconSet("4Arrows") },
        //        {"4ArrowsGray", GetIconSet("4ArrowsGray") },
        //        {"4RedToBlack", GetIconSet("4RedToBlack") },
        //        {"4Rating", GetIconSet("4Rating") },
        //        {"4TrafficLights", GetIconSet("4TrafficLights") },
        //        {"5Rating", GetIconSet("5Rating") },
        //        {"5Quarters", GetIconSet("5Quarters") },
        //        {"5Boxes", GetIconSet("5Boxes") },
        //        {"NoIcons", GetIconSet("NoIcons") },
        //    };

        //static readonly Dictionary<string, int> _iconStringSetDictionary = new Dictionary<string, int>
        //    {
        //     { "3Arrows" , 0 },
        //     { "3ArrowsGray" , 1 },
        //     { "3Flags" , 2 },
        //     { "3TrafficLights1" , 3 } ,
        //     { "3TrafficLights2" , 4 },
        //     { "3Signs" , 5 },
        //     { "3Symbols" , 6 },
        //     { "3Symbols2" , 7 },
        //     { "3Stars" , 8 },
        //     { "3Triangles" , 9 },
        //     { "4Arrows" , 10 },
        //     { "4ArrowsGray" , 11 },
        //     { "4RedToBlack" , 12 },
        //     { "4Rating" , 13 },
        //     { "4TrafficLights" , 14 },
        //     { "5Rating" , 15 },
        //     { "5Quarters" , 16 },
        //     { "5Boxes" , 17 },
        //     { "NoIcons" , 18},
        //    };

        //static readonly Dictionary<string, eExcelconditionalFormattingCustomIcon[]> _iconStringSetDictionarySets = new Dictionary<string, eExcelconditionalFormattingCustomIcon[]>
        //    {
        //     { "3Arrows" , {eExcelconditionalFormattingCustomIcon.RedDownArrow} },
        //     { "3ArrowsGray" , 1 },
        //     { "3Flags" , 2 },
        //     { "3TrafficLights1" , 3 } ,
        //     { "3TrafficLights2" , 4 },
        //     { "3Signs" , 5 },
        //     { "3Symbols" , 6 },
        //     { "3Symbols2" , 7 },
        //     { "3Stars" , 8 },
        //     { "3Triangles" , 9 },
        //     { "4Arrows" , 10 },
        //     { "4ArrowsGray" , 11 },
        //     { "4RedToBlack" , 12 },
        //     { "4Rating" , 13 },
        //     { "4TrafficLights" , 14 },
        //     { "5Rating" , 15 },
        //     { "5Quarters" , 16 },
        //     { "5Boxes" , 17 },
        //     { "NoIcons" , 18},
        //    };



        //internal static eExcelconditionalFormattingCustomIcon[] GetIconAtIndicies(string set, int[] indicies)
        //{
        //    var setValue = _iconStringSetDictionary[set];
        //    var iconValueBase = setValue << 4;

        //    eExcelconditionalFormattingCustomIcon[] retArr = new eExcelconditionalFormattingCustomIcon[indicies.Length];

        //    for (int i = 0; i < retArr.Length; i++)
        //    {
        //        retArr[i] = ConvertIntIdToEnum(iconValueBase, indicies[i]);
        //    }

        //    return retArr;
        //}

        //internal static eExcelconditionalFormattingCustomIcon[] GetIconAtIndicies(int iconValueBase, int[] indicies)
        //{
        //    eExcelconditionalFormattingCustomIcon[] retArr = new eExcelconditionalFormattingCustomIcon[indicies.Length];

        //    for (int i = 0; i < retArr.Length; i++)
        //    {
        //        retArr[i] = ConvertIntIdToEnum(iconValueBase, indicies[i]);
        //    }

        //    return retArr;
        //}

        //internal static eExcelconditionalFormattingCustomIcon GetIconAtIndex(string set, int index)
        //{
        //    var setValue = _iconStringSetDictionary[set];
        //    var iconValueBase = setValue << 4;

        //    return ConvertIntIdToEnum(iconValueBase, index);
        //}

        //internal static eExcelconditionalFormattingCustomIcon GetIcon(string set, int index, int iconValueBase = -1)
        //{
        //    if(iconValueBase == -1)
        //    {
        //        var setValue = _iconStringSetDictionary[set];
        //        iconValueBase = setValue << 4;
        //    }

        //    int iconValue = iconValueBase + index;

        //    return ConvertIntIdToEnum(iconValueBase, index);
        //}

        //private static eExcelconditionalFormattingCustomIcon ConvertIntIdToEnum(int iconValueBase, int index)
        //{
        //    var iconValue = iconValueBase + index;

        //    if (Enum.IsDefined(typeof(eExcelconditionalFormattingCustomIcon), iconValue) == false)
        //    {
        //        if (iconValue == 82)
        //        {
        //            iconValue = (int)eExcelconditionalFormattingCustomIcon.GreenCircle;
        //        }
        //        if (iconValueBase >= 160)
        //        {
        //            if (iconValue > 160 && index > 2)
        //            {
        //                iconValue -= (160 + index - 2);
        //            }
        //            else
        //            {
        //                iconValue -= 160;
        //            }
        //        }

        //        if(Enum.IsDefined(typeof(eExcelconditionalFormattingCustomIcon), iconValue) == false)
        //        {
        //            throw new NotImplementedException($"The id:{iconValue}");
        //        }
        //    }
            ////Special case
            //if (iconValue == 82)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.GreenCircle;
            //}else if(iconValue == 160)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.RedDownArrow;
            //}else if(iconValue == 163)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.GreenUpArrow;
            //}
            //else if(iconValue == 176)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.GrayDownArrow;
            //}
            //else if(iconValue == 179)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.GrayUpArrow;
            //}
            //else if(iconValue == 260)
            //{
            //    iconValue = (int)eExcelconditionalFormattingCustomIcon.BlackCircle;
            //}
            //else if (iconValue > 224 && iconValue < 228)
            //{
            //    //Move to circles
            //    var val = iconValue - 225;
            //    iconValue = 47 + val;
            //}

        //    return (eExcelconditionalFormattingCustomIcon)iconValue;
        //}


        //internal static eExcelconditionalFormattingCustomIcon[] GetIconSet(string set)
        //{
        //    int size = 3;
        //    if(set[0] == '4')
        //    {
        //        size = 4;
        //    }
        //    else if(set[0] == '5')
        //    {
        //        size = 5;
        //    }
        //    else if (set[0] == 'N')
        //    {
        //        size = 1;
        //    }

        //    //eExcelconditionalFormattingCustomIcon[] retArr = new eExcelconditionalFormattingCustomIcon[size];

        //    //var setValue = _iconStringSetDictionary[set];
        //    //var iconValueBase = setValue << 4;

        //    //for (int i = 0; i < retArr.Length; i++)
        //    //{
        //    //    retArr[i] = ConvertIntIdToEnum(iconValueBase + i);
        //    //}

        //    //return retArr;

        //    var arr = new int[size];

        //    for (int i = 0; i < arr.Length; i++)
        //    {
        //        arr[i] = i;
        //    }

        //    return GetIconAtIndicies(set, arr);
        //}

        internal static eExcelconditionalFormattingCustomIcon[] GetIconSet(string set)
        {
            var list = IconSets[set];
            return list.ToArray();
        }

        internal static eExcelconditionalFormattingCustomIcon GetIconAtIndex(string set, int index)
        {
            return IconSets[set][index];
        }



        //internal static virtual string GetCustomIconStringValue()
        //{
        //    if (CustomIcon != null)
        //    {
        //        int customIconId = (int)CustomIcon;

        //        var iconSetId = customIconId >> 4;

        //        return _iconStringSetDictionary[iconSetId];
        //    }

        //    throw new NotImplementedException($"Cannot get custom icon {CustomIcon} of {this} ");
        //}

        //internal static int GetCustomIconIndex()
        //{
        //    if (CustomIcon != null)
        //    {
        //        return (int)CustomIcon & 0xf;
        //    }

        //    return -1;
        //}
    }
}
