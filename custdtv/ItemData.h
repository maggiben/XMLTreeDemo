/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 1999 Microsoft Corporation.  All Rights Reserved.

   ---------------------------------------------------------------------
   
   Krishna Kotipalli, Microsoft Developer Support, SHell/UI

**************************************************************************/

// ItemData.h

#ifndef _ITEMDATA_H
#define _ITEMDATA_H

typedef enum _itemType
{
   ITEM_ROOT,
   ITEM_COLORS,
   ITEM_FONTS,
   ITEM_UNKNOWN
}ITEMTYPE;

typedef union _itemData
{
   struct
   {
      COLORREF clr;
      int clrTag; // index into the static _szClrTag array
   }clrData;

   struct
   {
      HFONT hFont;   // delete this on control destruction (caching this efficiency)
      int fontTag;   // index into the dynamically created _szFontTag array
   }fontData;
}ITEMDATA;

struct CItemData
{
   CItemData(ITEMTYPE itemType)
   {
      m_itemType = itemType;

      m_itemData.clrData.clr= 0;
      m_itemData.clrData.clrTag= 0;

      m_itemData.fontData.hFont = 0;
      m_itemData.fontData.fontTag = 0;
   }
   ~CItemData()
   {
      if (m_itemType == ITEM_FONTS)
      {
         DeleteObject(m_itemData.fontData.hFont);
      }
   }
   ITEMTYPE m_itemType;
   ITEMDATA m_itemData;
};

#endif //_ITEMDATA_H

