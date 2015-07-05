/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 1999 Microsoft Corporation.  All Rights Reserved.

   ---------------------------------------------------------------------
   
   Krishna Kotipalli, Microsoft Developer Support, SHell/UI

**************************************************************************/

// CustDraw.cpp : Illustrates custom draw in a TreeView control
//                Requires Version 4.70 of the ComCtl32.dll

#include "stdafx.h"
#include <commctrl.h>
#include <tchar.h>

#include "CustDTv.h"
#include "ItemData.h"
#include "xcolors.h"

long handleCustomDraw(HWND hWndTreeView, LPNMTVCUSTOMDRAW pNMTVCD)
{
   if (pNMTVCD==NULL)
   {
      return -1;
   }
   switch (pNMTVCD->nmcd.dwDrawStage)
   { 
      case CDDS_PREPAINT:
         return (CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYITEMDRAW);
      case CDDS_ITEMPREPAINT:
         {

            HTREEITEM hItem = (HTREEITEM) pNMTVCD->nmcd.dwItemSpec;
            TVITEM tvi = { 0 };
            tvi.mask = TVIF_HANDLE | TVIF_PARAM;
            tvi.hItem = hItem;
            TreeView_GetItem(hWndTreeView, &tvi);
            if (tvi.lParam)
            {
               CItemData *pItemData = (CItemData *) tvi.lParam;
               if (pItemData)
               {
                  if (pItemData->m_itemType == ITEM_COLORS)
                  {
                     if ((pNMTVCD->nmcd.uItemState & CDIS_FOCUS) == CDIS_FOCUS)
                     {

                        HFONT hFont = (HFONT) ::SendMessage(hWndTreeView, WM_GETFONT, 0, 0);
                        
                        LOGFONT lf = { 0 };
                        GetObject(hFont, sizeof(LOGFONT), &lf);

                        lf.lfWeight |= FW_BOLD;
                        
                        HFONT hFontBold = CreateFontIndirect(&lf);
                        HFONT hOldFont = (HFONT) SelectObject(pNMTVCD->nmcd.hdc, hFontBold);

                        DeleteObject(hFontBold);
                     }
                     else
                     {
                        pNMTVCD->clrText = pItemData->m_itemData.clrData.clr;
                     }
                  }
                  else if (pItemData->m_itemType == ITEM_FONTS)
                  {
                     SelectObject(pNMTVCD->nmcd.hdc, pItemData->m_itemData.fontData.hFont);
                  }
               }
            }
            return (CDRF_NOTIFYPOSTPAINT | CDRF_NEWFONT);
         }
      case CDDS_ITEMPOSTPAINT:
         {
            RECT rc;
            TreeView_GetItemRect(hWndTreeView, (HTREEITEM) pNMTVCD->nmcd.dwItemSpec, &rc, 1);

            int temp = rc.left;
            //rc.top += 1;
            rc.left = rc.right + GAP_SIZE;
            rc.right += temp + rc.right + GAP_SIZE;

            
            TCHAR szFace[LF_FACESIZE];
            HTREEITEM hItem = (HTREEITEM) pNMTVCD->nmcd.dwItemSpec;
            TVITEM tvi = { 0 };
            tvi.mask = TVIF_HANDLE | TVIF_PARAM | TVIF_TEXT;
            tvi.hItem = hItem;
            tvi.pszText = szFace;
            tvi.cchTextMax = sizeof(szFace)/sizeof(TCHAR);

            TreeView_GetItem(hWndTreeView, &tvi);
            if (tvi.lParam)
            {
               CItemData *pItemData = (CItemData *) tvi.lParam;
               if (pItemData)
               {
                  if (pItemData->m_itemType == ITEM_COLORS)
                  {
                     if ((pNMTVCD->nmcd.uItemState & CDIS_FOCUS) == CDIS_FOCUS)
                     {
                        // Draw the Color Text
                        int nOldClr = SetTextColor(pNMTVCD->nmcd.hdc, RGB(0,0,255));
                        DrawText(pNMTVCD->nmcd.hdc, _szClrTag[pItemData->m_itemData.clrData.clrTag], -1, &rc, DT_LEFT);
                        SetTextColor(pNMTVCD->nmcd.hdc, nOldClr);
                     }
                     else
                     {
                        // Erase the Color Text drawn previously
                        FillRect(pNMTVCD->nmcd.hdc, &rc, (HBRUSH) (COLOR_WINDOW+1));
                     }
                  }
                  else if (pItemData->m_itemType == ITEM_FONTS)
                  {
                  }
               }
            }
            return CDRF_DODEFAULT;
         }
         break;
      default:
         break;
   }
   return 0;
}

//EOF
