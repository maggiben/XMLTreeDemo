/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 1999 Microsoft Corporation.  All Rights Reserved.

   ---------------------------------------------------------------------
   
   Krishna Kotipalli, Microsoft Developer Support, SHell/UI

**************************************************************************/

// xcolors.h

#ifndef _XCOLORS_H_
#define _XCOLORS_H_

#define MAX_COLORS 99

static TCHAR *_szClrTag[MAX_COLORS] = 
{
   _T("White"),
   _T("Red"),
   _T("Green"),
   _T("Blue"),
   _T("Magenta"),
   _T("Cyan"),
   _T("Yellow"),
   _T("Black"),
   _T("Aquamarine"),
   _T("Baker's Chocolate"),
   _T("Blue Violet"),
   _T("Brass"),
   _T("Bright Gold"),
   _T("Brown"),
   _T("Bronze"),
   _T("Bronze II"),
   _T("Cadet Blue"),
   _T("Cool Copper"),
   _T("Copper"),
   _T("Coral"),
   _T("Corn Flower Blue"),
   _T("Dark Brown"),
   _T("Dark Green"),
   _T("Dark Green Copper"),
   _T("Dark Olive Green"),
   _T("Dark Orchid"),
   _T("Dark Purple"),
   _T("Dark Slate Blue"),
   _T("Dark Slate Grey"),
   _T("Dark Tan"),
   _T("Dark Turquoise"),
   _T("Dark Wood"),
   _T("Dim Grey"),
   _T("Dusty Rose"),
   _T("Feldspar"),
   _T("Firebrick"),
   _T("Forest Green"),
   _T("Gold"),
   _T("Goldenrod"),
   _T("Grey"),
   _T("Green Copper"),
   _T("Green Yellow"),
   _T("Hunter Green"),
   _T("Indian Red"),
   _T("Khaki"),
   _T("Light Blue"),
   _T("Light Grey"),
   _T("Light Steel Blue"),
   _T("Light Wood"),
   _T("Lime Green"),
   _T("Mandarian Orange"),
   _T("Maroon"),
   _T("Medium Aquamarine"),
   _T("Medium Blue"),
   _T("Medium Forest Green"),
   _T("Medium Goldenrod"),
   _T("Medium Orchid"),
   _T("Medium Sea Green"),
   _T("Medium Slate Blue"),
   _T("Medium Spring Green"),
   _T("Medium Turquoise"),
   _T("Medium Violet Red"),
   _T("Medium Wood"),
   _T("Midnight Blue"),
   _T("Navy Blue"),
   _T("Neon Blue"),
   _T("Neon Pink"),
   _T("New Midnight Blue"),
   _T("New Tan"),
   _T("Old Gold"),
   _T("Orange"),
   _T("Orange Red"),
   _T("Orchid"),
   _T("Pale Green"),
   _T("Pink"),
   _T("Plum"),
   _T("Quartz"),
   _T("Rich Blue"),
   _T("Salmon"),
   _T("Scarlet"),
   _T("Sea Green"),
   _T("Semi-Sweet Chocolate"),
   _T("Sienna"),
   _T("Silver"),
   _T("Sky Blue"),
   _T("Slate Blue"),
   _T("Spicy Pink"),
   _T("Spring Green"),
   _T("Steel Blue"),
   _T("Summer Sky"),
   _T("Tan"),
   _T("Thistle"),
   _T("Turquoise"),
   _T("Very Dark Brown"),
   _T("Very Light Grey"),
   _T("Violet"),
   _T("Violet Red"),
   _T("Wheat"),
   _T("Yellow Green")
};

static COLORREF _clraText[MAX_COLORS] = 
{
   RGB(0xFF,0xFF,0xFF),   // White
   RGB(0xFF,0x00,0x00),   // Red
   RGB(0x00,0xFF,0x00),   // Green
   RGB(0x00,0x00,0xFF),   // Blue
   RGB(0xFF,0x00,0xFF ),   // Magenta
   RGB(0x00,0xFF,0xFF),   // Cyan
   RGB(0xFF,0xFF,0x00),   // Yellow
   RGB(0x00,0x00,0x00),   // Black
   RGB(0x70,0xDB,0x93),   // Aquamarine
   RGB(0x5C,0x33,0x17),   // Baker's Chocolate
   RGB(0x9F,0x5F,0x9F),   // Blue Violet
   RGB(0xB5,0xA6,0x42),   // Brass
   RGB(0xD9,0xD9,0x19),   // Bright Gold
   RGB(0xA6,0x2A,0x2A),   // Brown
   RGB(0x8C,0x78,0x53),   // Bronze
   RGB(0xA6,0x7D,0x3D),   // Bronze II
   RGB(0x5F,0x9F,0x9F),   // Cadet Blue
   RGB(0xD9,0x87,0x19),   // Cool Copper
   RGB(0xB8,0x73,0x33 ),   // Copper
   RGB(0xFF,0x7F,0x00),   // Coral
   RGB(0x42,0x42,0x6F),   // Corn Flower Blue
   RGB(0x5C,0x40,0x33 ),   // Dark Brown
   RGB(0x2F,0x4F,0x2F),   // Dark Green
   RGB(0x4A,0x76,0x6E),   // Dark Green Copper
   RGB(0x4F,0x4F,0x2F),   // Dark Olive Green
   RGB(0x99,0x32,0xCD),   // Dark Orchid
   RGB(0x87,0x1F,0x78),   // Dark Purple
   RGB(0x6B,0x23,0x8E),   // Dark Slate Blue
   RGB(0x2F,0x4F,0x4F),   // Dark Slate Grey
   RGB(0x97,0x69,0x4F),   // Dark Tan
   RGB(0x70,0x93,0xDB),   // Dark Turquoise
   RGB(0x85,0x5E,0x42),   // Dark Wood
   RGB(0x54,0x54,0x54),   // Dim Grey
   RGB(0x85,0x63,0x63),   // Dusty Rose
   RGB(0xD1,0x92,0x75),   // Feldspar
   RGB(0x8E,0x23,0x23),   // Firebrick
   RGB(0x23,0x8E,0x23),   // Forest Green
   RGB(0xCD,0x7F,0x32 ),   // Gold
   RGB(0xDB,0xDB,0x70),   // Goldenrod
   RGB(0xC0,0xC0,0xC0),   // Grey
   RGB(0x52,0x7F,0x76),   // Green Copper
   RGB(0x93,0xDB,0x70),   // Green Yellow
   RGB(0x21,0x5E,0x21),   // Hunter Green
   RGB(0x4E,0x2F,0x2F),   // Indian Red
   RGB(0x9F,0x9F,0x5F),   // Khaki
   RGB(0xC0,0xD9,0xD9),   // Light Blue
   RGB(0xA8,0xA8,0xA8),   // Light Grey
   RGB(0x8F,0x8F,0xBD),   // Light Steel Blue
   RGB(0xE9,0xC2,0xA6),   // Light Wood
   RGB(0x32,0xCD,0x32),   // Lime Green
   RGB(0xE4,0x78,0x33),   // Mandarian Orange
   RGB(0x8E,0x23,0x6B),   // Maroon
   RGB(0x32,0xCD,0x99),   // Medium Aquamarine
   RGB(0x32,0x32,0xCD),   // Medium Blue
   RGB(0x6B,0x8E,0x23),   // Medium Forest Green
   RGB(0xEA,0xEA,0xAE),   // Medium Goldenrod
   RGB(0x93,0x70,0xDB),   // Medium Orchid
   RGB(0x42,0x6F,0x42),   // Medium Sea Green
   RGB(0x7F,0x00,0xFF),   // Medium Slate Blue
   RGB(0x7F,0xFF,0x00),   // Medium Spring Green
   RGB(0x70,0xDB,0xDB),   // Medium Turquoise
   RGB(0xDB,0x70,0x93),   // Medium Violet Red
   RGB(0xA6,0x80,0x64),   // Medium Wood
   RGB(0x2F,0x2F,0x4F),   // Midnight Blue
   RGB(0x23,0x23,0x8E),   // Navy Blue
   RGB(0x4D,0x4D,0xFF),   // Neon Blue
   RGB(0xFF,0x6E,0xC7),   // Neon Pink
   RGB(0x00,0x00,0x9C),   // New Midnight Blue
   RGB(0xEB,0xC7,0x9E),   // New Tan
   RGB(0xCF,0xB5,0x3B),   // Old Gold
   RGB(0xFF,0x7F,0x00),   // Orange
   RGB(0xFF,0x24,0x00),   // Orange Red
   RGB(0xDB,0x70,0xDB),   // Orchid
   RGB(0x8F,0xBC,0x8F),   // Pale Green
   RGB(0xBC,0x8F,0x8F),   // Pink
   RGB(0xEA,0xAD,0xEA),   // Plum
   RGB(0xD9,0xD9,0xF3),   // Quartz
   RGB(0x59,0x59,0xAB),   // Rich Blue
   RGB(0x6F,0x42,0x42),   // Salmon
   RGB(0x8C,0x17,0x17),   // Scarlet
   RGB(0x23,0x8E,0x68),   // Sea Green
   RGB(0x6B,0x42,0x26),   // Semi-Sweet Chocolate
   RGB(0x8E,0x6B,0x23),   // Sienna
   RGB(0xE6,0xE8,0xFA),   // Silver
   RGB(0x32,0x99,0xCC),   // Sky Blue
   RGB(0x00,0x7F,0xFF),   // Slate Blue
   RGB(0xFF,0x1C,0xAE),   // Spicy Pink
   RGB(0x00,0xFF,0x7F),   // Spring Green
   RGB(0x23,0x6B,0x8E),   // Steel Blue
   RGB(0x38,0xB0,0xDE),   // Summer Sky
   RGB(0xDB,0x93,0x70),   // Tan
   RGB(0xD8,0xBF,0xD8),   // Thistle
   RGB(0xAD,0xEA,0xEA),   // Turquoise
   RGB(0x5C,0x40,0x33),   // Very Dark Brown
   RGB(0xCD,0xCD,0xCD),   // Very Light Grey
   RGB(0x4F,0x2F,0x4F),   // Violet
   RGB(0xCC,0x32,0x99),   // Violet Red
   RGB(0xD8,0xD8,0xBF),   // Wheat
   RGB(0x99,0xCC,0x32)      // Yellow Green
};

static TCHAR *_szItem[MAX_COLORS] = 
{
   _T("0xFFFFFF   "),
   _T("0xFF0000   "),
   _T("0x00FF00   "),
   _T("0x0000FF   "),
   _T("0xFF00FF"   ), 
   _T("0x00FFFF   "),
   _T("0xFFFF00   "),
   _T("0x000000   "),
   _T("0x70DB93   "),
   _T("0x5C3317   "),
   _T("0x9F5F9F   "),
   _T("0xB5A642   "),
   _T("0xD9D919   "),
   _T("0xA62A2A   "),
   _T("0x8C7853   "),
   _T("0xA67D3D   "),
   _T("0x5F9F9F   "),
   _T("0xD98719   "),
   _T("0xB87333   "), 
   _T("0xFF7F00   "),
   _T("0x42426F   "),
   _T("0x5C4033   "), 
   _T("0x2F4F2F   "),
   _T("0x4A766E   "),
   _T("0x4F4F2F   "),
   _T("0x9932CD   "),
   _T("0x871F78   "),
   _T("0x6B238E   "),
   _T("0x2F4F4F   "),
   _T("0x97694F   "),
   _T("0x7093DB   "),
   _T("0x855E42   "),
   _T("0x545454   "),
   _T("0x856363   "),
   _T("0xD19275   "),
   _T("0x8E2323   "),
   _T("0x238E23   "),
   _T("0xCD7F32   "), 
   _T("0xDBDB70   "),
   _T("0xC0C0C0   "),
   _T("0x527F76   "),
   _T("0x93DB70   "),
   _T("0x215E21   "),
   _T("0x4E2F2F   "),
   _T("0x9F9F5F   "),
   _T("0xC0D9D9   "),
   _T("0xA8A8A8   "),
   _T("0x8F8FBD   "),
   _T("0xE9C2A6   "),
   _T("0x32CD32   "),
   _T("0xE47833   "),
   _T("0x8E236B   "),
   _T("0x32CD99   "),
   _T("0x3232CD   "),
   _T("0x6B8E23   "),
   _T("0xEAEAAE   "),
   _T("0x9370DB   "),
   _T("0x426F42   "),
   _T("0x7F00FF   "),
   _T("0x7FFF00   "),
   _T("0x70DBDB   "),
   _T("0xDB7093   "),
   _T("0xA68064   "),
   _T("0x2F2F4F   "),
   _T("0x23238E   "),
   _T("0x4D4DFF   "),
   _T("0xFF6EC7   "),
   _T("0x00009C   "),
   _T("0xEBC79E   "),
   _T("0xCFB53B   "),
   _T("0xFF7F00   "),
   _T("0xFF2400   "),
   _T("0xDB70DB   "),
   _T("0x8FBC8F   "),
   _T("0xBC8F8F   "),
   _T("0xEAADEA   "),
   _T("0xD9D9F3   "),
   _T("0x5959AB   "),
   _T("0x6F4242   "),
   _T("0x8C1717   "),
   _T("0x238E68   "),
   _T("0x6B4226   "),
   _T("0x8E6B23   "),
   _T("0xE6E8FA   "),
   _T("0x3299CC   "),
   _T("0x007FFF   "),
   _T("0xFF1CAE   "),
   _T("0x00FF7F   "),
   _T("0x236B8E   "),
   _T("0x38B0DE   "),
   _T("0xDB9370   "),
   _T("0xD8BFD8   "),
   _T("0xADEAEA   "),
   _T("0x5C4033   "),
   _T("0xCDCDCD   "),
   _T("0x4F2F4F   "),
   _T("0xCC3299   "),
   _T("0xD8D8BF   "),
   _T("0x99CC32   ")
};

#endif // _XCOLORS_H_

