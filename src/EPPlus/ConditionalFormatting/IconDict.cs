﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class IconDict
    {

        //internal static readonly Dictionary<string, List<eExcelconditionalFormattingCustomIcon>> IconSets
        //    = new Dictionary<string, List<eExcelconditionalFormattingCustomIcon>>
        //    {
        //        {"3Arrows", ThreeArrows },
        //        {"3ArrowsGray", ThreeArrowsGray },
        //    };


        //static readonly List<eExcelconditionalFormattingCustomIcon> ThreeArrows =
        //    new List<eExcelconditionalFormattingCustomIcon>()
        //    { eExcelconditionalFormattingCustomIcon.RedDownArrow, eExcelconditionalFormattingCustomIcon.YellowSideArrow, eExcelconditionalFormattingCustomIcon.GreenUpArrow};

        //static readonly List<eExcelconditionalFormattingCustomIcon> ThreeArrowsGray =
        //    new List<eExcelconditionalFormattingCustomIcon>()
        //    { eExcelconditionalFormattingCustomIcon.GrayDownArrow, eExcelconditionalFormattingCustomIcon.YellowSideArrow, eExcelconditionalFormattingCustomIcon.GrayUpArrow};

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

        static readonly Dictionary<string, int> _iconStringSetDictionary = new Dictionary<string, int>
            {
             { "3Arrows" , 0 },
             { "3ArrowsGray" , 1 },
             { "3Flags" , 2 },
             { "3TrafficLights1" , 3 } ,
             { "3TrafficLights2" , 4 },
             { "3Signs" , 5 },
             { "3Symbols" , 6 },
             { "3Symbols2" , 7 },
             { "3Stars" , 8 },
             { "3Triangles" , 9 },
             { "4Arrows" , 10 },
             { "4ArrowsGray" , 11 },
             { "4RedToBlack" , 12 },
             { "4Rating" , 13 },
             { "4TrafficLights" , 14 },
             { "5Rating" , 15 },
             { "5Quarters" , 16 },
             { "5Boxes" , 17 },
             { "NoIcons" , 18},
            };

        internal static eExcelconditionalFormattingCustomIcon[] GetIconSet(string set)
        {
            int size = 3;
            if(set[0] == '4')
            {
                size = 4;
            }
            else if(set[0] == '5')
            {
                size = 5;
            }
            else if (set[0] == 'N')
            {
                size = 1;
            }

            eExcelconditionalFormattingCustomIcon[] retArr = new eExcelconditionalFormattingCustomIcon[size];

            for (int i = 0; i < retArr.Length; i++)
            {
                var setValue = _iconStringSetDictionary[set];
                int iconValue = setValue << 4;
                iconValue += i;
                retArr[i] = (eExcelconditionalFormattingCustomIcon)iconValue;
            }

            return retArr;
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
