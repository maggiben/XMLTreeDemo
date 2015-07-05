/**************************************************************************
   THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
   ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
   PARTICULAR PURPOSE.

   Copyright 1999 Microsoft Corporation.  All Rights Reserved.

   ---------------------------------------------------------------------
   
   Krishna Kotipalli, Microsoft Developer Support, SHell/UI

**************************************************************************/

// CustDTv.h
#ifndef _CUSTDTV_H_
#define _CUSTDTV_H_

#define GAP_SIZE         5

#define SZ_COLORS         _T("Some Colors and their (HTML) names")
#define SZ_FONTS         _T("Fonts installed in the system")

static BYTE _baCharSets[] = 
{
   ANSI_CHARSET,
   BALTIC_CHARSET,
   CHINESEBIG5_CHARSET,
   DEFAULT_CHARSET,
   EASTEUROPE_CHARSET,
   GB2312_CHARSET,
   GREEK_CHARSET,
   HANGUL_CHARSET,
   MAC_CHARSET,
   OEM_CHARSET,
   RUSSIAN_CHARSET,
   SHIFTJIS_CHARSET,
   SYMBOL_CHARSET,
   TURKISH_CHARSET,
   0xFF
};

static TCHAR *_szCharSets[] = 
{
   _T("ANSI_CHARSET"),
   _T("BALTIC_CHARSET"),
   _T("CHINESEBIG5_CHARSET"),
   _T("DEFAULT_CHARSET"),
   _T("EASTEUROPE_CHARSET"),
   _T("GB2312_CHARSET"),
   _T("GREEK_CHARSET"),
   _T("HANGUL_CHARSET"),
   _T("MAC_CHARSET"),
   _T("OEM_CHARSET"),
   _T("RUSSIAN_CHARSET"),
   _T("SHIFTJIS_CHARSET"),
   _T("SYMBOL_CHARSET"),
   _T("TURKISH_CHARSET"),
   _T("\0")
};

struct fontProcData
{
   HWND hWndTreeView;
   HTREEITEM hParentItem;
   BYTE lfCharSet;
   LOGFONT *plfTreeView;
};
int CALLBACK enumFontFamilyProc(ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, DWORD FontType, LPARAM lParam);

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);

void insertItems(HWND hWndTreeView);

long handleNotify(HWND hwndDlg, int nIDCtrl, LPNMHDR pNMHDR);
long handleCustomDraw(HWND hWndTreeView, LPNMTVCUSTOMDRAW pNMTVCD);
int getChilds(HWND hWnd, HTREEITEM hItem);
void destroyTreeItemData(HWND hWndTreeView);
void destroyItemData(HWND hWndTreeView, HTREEITEM hItem);

void enumFonts(HWND hWndTreeView, HTREEITEM hItem, BYTE lfCharSet);

#endif //_CUSTDTV_H_

