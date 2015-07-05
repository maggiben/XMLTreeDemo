/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 1999 Microsoft Corporation.  All Rights Reserved.

   ---------------------------------------------------------------------
   
   Krishna Kotipalli, Microsoft Developer Support, SHell/UI

**************************************************************************/

// CustDTv.cpp : Defines the entry point for the application.
//

#include "stdafx.h"

#include <commctrl.h>
#include <stdlib.h>
#include <tchar.h>

#include "resource.h"
#include "xcolors.h"
#include "itemdata.h"
#include "CustDTv.h"

#pragma comment(lib, "comctl32.lib")

int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
   INITCOMMONCONTROLSEX icc = { 0 };
   icc.dwSize = sizeof(INITCOMMONCONTROLSEX);
   icc.dwICC = ICC_TREEVIEW_CLASSES;

   BOOL bRet = InitCommonControlsEx(&icc);

   DialogBox(hInstance, MAKEINTRESOURCE(IDD_DIALOG), NULL, (DLGPROC) DialogProc);

   return 0;
}



INT_PTR CALLBACK DialogProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
   switch (uMsg)
   {
   case WM_COMMAND:
      switch (LOWORD(wParam))
      {
      case IDOK:
         EndDialog(hWndDlg, IDOK);
         break;
      case IDCANCEL:
         EndDialog(hWndDlg, IDCANCEL);
         break;
      }
      break;
   case WM_INITDIALOG:
      {
         HWND hWndTreeView = GetDlgItem(hWndDlg, IDC_TREE);
         SetFocus(hWndTreeView);

         insertItems(hWndTreeView);
      }
      return FALSE; // FALSE since we set focus to hWndTreeView
   case WM_DESTROY:
      {
         HWND hWndTreeView = GetDlgItem(hWndDlg, IDC_TREE);
         destroyTreeItemData(hWndTreeView);
      }
      break;
   case WM_NOTIFY:
      {
         long lResult = handleNotify(hWndDlg, (int) wParam, (LPNMHDR) lParam);
         
         // This is the way a dialog box can send back lresults.. 
         // See documentation for DialogProc for more information
         SetWindowLong(hWndDlg, DWL_MSGRESULT, lResult); 
         return TRUE;
      }
      break;
   }
   return FALSE;
}

void insertItems(HWND hWndTreeView)
{
   TVINSERTSTRUCT tvis = { 0 };

   CItemData *pClrItemData = new CItemData(ITEM_UNKNOWN);
   tvis.item.mask = TVIF_TEXT | TVIF_PARAM;
   tvis.item.lParam = (LPARAM) pClrItemData;
   tvis.item.pszText = SZ_COLORS;
   tvis.item.cchTextMax = _tcslen(SZ_COLORS);

   HTREEITEM hItem = TreeView_InsertItem(hWndTreeView, &tvis);
   TreeView_SelectItem(hWndTreeView, hItem);

   for (int i = 0; i < MAX_COLORS; i++)
   {
      CItemData *pItemData = new CItemData(ITEM_COLORS);

      pItemData->m_itemData.clrData.clr = _clraText[i];
      pItemData->m_itemData.clrData.clrTag = i;
      
      tvis.item.mask = TVIF_TEXT | TVIF_PARAM;
      tvis.item.lParam = (LPARAM) pItemData;
      tvis.hParent = hItem;
      tvis.item.pszText = _szItem[i];
      tvis.item.cchTextMax = _tcslen(_szItem[i]);
      
      TreeView_InsertItem(hWndTreeView, &tvis);
   }

   CItemData *pFontItemData = new CItemData(ITEM_UNKNOWN);

   tvis.item.mask = TVIF_TEXT | TVIF_PARAM;
   tvis.item.lParam = (LPARAM) pFontItemData;
   tvis.hParent = NULL;
   tvis.item.pszText = SZ_FONTS;
   tvis.item.cchTextMax = _tcslen(SZ_FONTS);
   
   hItem = TreeView_InsertItem(hWndTreeView, &tvis);

   tvis.item.mask = TVIF_TEXT;
   tvis.hParent = hItem;
   
   for (i = 0; _baCharSets[i] != 0xFF; i++)
   {
      tvis.item.pszText = _szCharSets[i];
      tvis.item.cchTextMax = _tcslen(_szCharSets[i]);
      hItem = TreeView_InsertItem(hWndTreeView, &tvis);

      enumFonts(hWndTreeView, hItem, _baCharSets[i]);
   }
}
         
void destroyTreeItemData(HWND hWndTreeView)
{
   HTREEITEM hRoot = TreeView_GetRoot(hWndTreeView);
   destroyItemData(hWndTreeView, hRoot);
   HTREEITEM hNextItem = TreeView_GetNextSibling(hWndTreeView, hRoot);
   while (hNextItem)
   {
      destroyItemData(hWndTreeView, hNextItem);
      hNextItem = TreeView_GetNextSibling(hWndTreeView, hNextItem);
   }
}

void destroyItemData(HWND hWndTreeView, HTREEITEM hItem)
{
   if (hItem)
   {
      TVITEM tvi = { 0 };
      tvi.mask = TVIF_HANDLE | TVIF_PARAM;
      tvi.hItem = hItem;
      TreeView_GetItem(hWndTreeView, &tvi);
      if (tvi.lParam)
      {
         CItemData *pItemData = (CItemData *) tvi.lParam;
         if (pItemData)
         {
            delete pItemData;
         }
      }
      HTREEITEM hChild = TreeView_GetChild(hWndTreeView, hItem);
      while (hChild)
      {
         destroyItemData(hWndTreeView, hChild);
         hChild = TreeView_GetNextSibling(hWndTreeView, hChild);
      }
   }
}

long handleNotify(HWND hWndDlg, int nIDCtrl, LPNMHDR pNMHDR)
{
   if (nIDCtrl == IDC_TREE && pNMHDR)
   {
      if (pNMHDR->code == NM_CUSTOMDRAW)
      {
         LPNMTVCUSTOMDRAW pNMTVCD = (LPNMTVCUSTOMDRAW) pNMHDR;
         HWND hWndTreeView = pNMHDR->hwndFrom;

         return handleCustomDraw(hWndTreeView, pNMTVCD);
      }
   }
   return 0;
}

int getChilds(HWND hWnd, HTREEITEM hItem)
{
   int nItems = 0;
   HTREEITEM hChild = TreeView_GetChild(hWnd, hItem);
   while (hChild)
   {
      nItems++;
      hChild = TreeView_GetNextSibling(hWnd, hChild);
   }
   return nItems;
}

void enumFonts(HWND hWndTreeView, HTREEITEM hItem, BYTE lfCharSet)
{
   HFONT hFont = (HFONT) ::SendMessage(hWndTreeView, WM_GETFONT, 0, 0);
   LOGFONT lfTreeView = { 0 };
   GetObject(hFont, sizeof(LOGFONT), &lfTreeView);

   HDC hDC = ::GetDC(NULL);

   fontProcData data = { hWndTreeView, hItem, lfCharSet, &lfTreeView };

   LOGFONT lf = { 0 };
   lf.lfCharSet = lfCharSet;
   _tcscpy(lf.lfFaceName, _T(""));

   EnumFontFamiliesEx(hDC, &lf, (FONTENUMPROC) enumFontFamilyProc, (LPARAM) &data, 0);
   
   ReleaseDC(NULL, hDC);
}

int CALLBACK enumFontFamilyProc(ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, DWORD FontType, LPARAM lParam)
{
   if (lpelfe && lParam)
   {
      fontProcData *pData  = (fontProcData *) lParam;

      if ((lpelfe->elfLogFont.lfCharSet & pData->lfCharSet) == lpelfe->elfLogFont.lfCharSet)
      {
         TVINSERTSTRUCT tvis = { 0 };

         CItemData *pItemData = new CItemData(ITEM_FONTS);

         lpelfe->elfLogFont.lfWidth = pData->plfTreeView->lfWidth;
         lpelfe->elfLogFont.lfHeight = pData->plfTreeView->lfHeight;

         pItemData->m_itemData.fontData.hFont = CreateFontIndirect(&lpelfe->elfLogFont);
         tvis.item.mask = TVIF_TEXT | TVIF_PARAM;
         tvis.item.lParam = (LPARAM) pItemData;
         tvis.hParent = pData->hParentItem;
         TCHAR szFontName[MAX_PATH];

         // 100 spaces to make sure that any item is not clipped, since
         // each item is drawn in a different font of its own...
         wsprintf(szFontName, _T("%-100s"), lpelfe->elfFullName); 
         
         tvis.item.pszText = szFontName;
         tvis.item.cchTextMax = _tcslen(szFontName);
         HTREEITEM hItem = TreeView_InsertItem(pData->hWndTreeView, &tvis);
      }
   }
   return 1;
}
