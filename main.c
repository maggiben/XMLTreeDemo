///////////////////////////////////////////////////////////////////////////////////
// File Name:           main.c                                                   //
// Author:              Benjamin Maggi                                           //
// Descripcion:         demostration app for XMLTree                             //
// Org. Date:           18/05/2009                                               //
// Last Modification:   20/05/2009                                               //
// Ver:                 0.0.3                                                    //
// compiler:            uses ansi-C / C99 tested with LCC & PellesC              //
// Author mail:         benjaminmaggi@gmail.com                                  //
// License:             GNU                                                      //
//                                                                               //
// License:                                                                      //
// This program is free software; you can redistribute it                        //
// and/or modify it under the terms of the GNU General Public                    //
// License as published by the Free Software Foundation;                         //
// either version 2 of the License, or (at your option) any                      //
// later version.                                                                //
//                                                                               //
// This program is distributed in the hope that it will be useful,               //
// but WITHOUT ANY WARRANTY; without even the implied warranty of                //
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                 //
// GNU General Public License for more details.                                  //
//                                                                               //
// Best viewed if you set your editor TAB width to 4.                            //
//                                                                               //
// Link me with:                                                                 //
// kernel32.lib delayimp.lib kernel32.lib user32.lib gdi32.lib comctl32.lib      //
// comdlg32.lib advapi32.lib uuid.lib ole32.lib oleaut32.lib shell32.lib         //
// oldnames.lib urlmon.lib                                                       //                                                  //
//                                                                               //
// Rev.2                                                                         //
//      * More Unicode compatible                                                //
//      * Save tree data to XML                                                  //
///////////////////////////////////////////////////////////////////////////////////


/* 
 * Either define WIN32_LEAN_AND_MEAN, or one or more of NOCRYPT,
 * NOSERVICE, NOMCX and NOIME, to decrease compile time (if you
 * don't need these defines -- see windows.h).
 */

#define _UNICODE
#define WIN32_LEAN_AND_MEAN
/* #define NOCRYPT */
/* #define NOSERVICE */
/* #define NOMCX */
/* #define NOIME */

#include <time.h>
#include <stdio.h>			// 
#include <windows.h>		// you know what this is for :)
#include <windowsx.h>
#include <commdlg.h>		// Open & Save file dialog
#include <tchar.h>          
#include <wchar.h>			// Wide Char
#include <ole2.h>           // Drag & Drop operations and more
#include <msxml.h>			// MSXML COM component header
#include <shlguid.h>
#include <shlobj.h>
#include <shldisp.h>
#include <shobjidl.h>
#include <oleauto.h>
#include <exdisp.h>
#include "main.h"

#include "XMLTree.h"

#define NELEMS(a)  (sizeof(a) / sizeof((a)[0]))

///////////////////////////////////////////////////////////////////////////////////
// Prototypes                                                                    //
///////////////////////////////////////////////////////////////////////////////////
static INT_PTR CALLBACK MainDlgProc(HWND, UINT, WPARAM, LPARAM);
static BOOL getBookData(HWND hwndDlg, IXMLDOMElement* node);
static BOOL setBookData(HWND hwndDlg, IXMLDOMElement* node);
int EnableGroupboxControls(HWND hWnd, BOOL bEnable);
int HiliteGroupboxControls(HWND hWnd, int id);
void CopyToClipboard(TCHAR *buf);

#define DLG_OPEN 0
#define DLG_SAVE 1
BOOL GetOpenSaveDlg(TCHAR *buffer, BYTE option);
///////////////////////////////////////////////////////////////////////////////////
// Global variables                                                              //
///////////////////////////////////////////////////////////////////////////////////
static HANDLE ghInstance;
const int MAX_GENRES = 5;
TCHAR *comboGenres[] = { _T("Computer"), _T("Fantasy"), _T("Horror"), _T("Romance"), _T("Science Fiction") };

/****************************************************************************
 *                                                                          *
 * Function: WinMain                                                        *
 *                                                                          *
 * Purpose : Initialize the application.  Register a window class,          *
 *           create and display the main window and enter the               *
 *           message loop.                                                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
    INITCOMMONCONTROLSEX icc;
    WNDCLASSEX wcx;

    ghInstance = hInstance;

    /* Initialize common controls. Also needed for MANIFEST's */
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_WIN95_CLASSES /*|ICC_COOL_CLASSES|ICC_DATE_CLASSES|ICC_PAGESCROLLER_CLASS|ICC_USEREX_CLASSES|... */;
    InitCommonControlsEx(&icc);

    /* Get system dialog information */
    wcx.cbSize = sizeof(wcx);
    if (!GetClassInfoEx(NULL, MAKEINTRESOURCE(32770), &wcx))
        return 0;

    /* Add our own stuff */
    wcx.hInstance = hInstance;
    wcx.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDR_ICO_MAIN));
    wcx.lpszClassName = _T("XMLTreeDClass");
    if (!RegisterClassEx(&wcx))
        return 0;

    /* The user interface is a modal dialog box */
    return DialogBox(hInstance, MAKEINTRESOURCE(DLG_MAIN), NULL, (DLGPROC)MainDlgProc);
}

/****************************************************************************
 *                                                                          *
 * Function: MainDlgProc                                                    *
 *                                                                          *
 * Purpose : Process messages for the Main dialog.                          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

static INT_PTR CALLBACK MainDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	static HWND selectedCtl = NULL;
    switch (uMsg)
    {
        case WM_INITDIALOG:
				SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_ADDSTRING, 0, (LPARAM)comboGenres[0]);
				SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_ADDSTRING, 0, (LPARAM)comboGenres[1]);
				SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_ADDSTRING, 0, (LPARAM)comboGenres[2]);
				SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_ADDSTRING, 0, (LPARAM)comboGenres[3]);
				SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_ADDSTRING, 0, (LPARAM)comboGenres[4]);
				
				EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
				PopulateTree(GetDlgItem(hwndDlg, IDC_TREE),_T("book.xml"));
            return TRUE;
		case WM_NOTIFY:
		{
			LPNMHDR param = (LPNMHDR)lParam;
			if(param->hwndFrom == GetDlgItem(hwndDlg, IDC_TREE))
			{
				///////////////////////////////////////////////
				// Reflect message our internal tree handler //
				// We will catch this messages to expand     //
				// the XML tree on demand:                   //
				// TVN_ITEMEXPANDING                         //
				// TVN_DELETEITEM                            //
				// If you'r going to use them also           //
				// Make sure you let our handler go first    //
				///////////////////////////////////////////////
				OnNotify(hwndDlg, uMsg, wParam, (NMHDR *)lParam);
				// Please do more stuff
				switch(param->code)
				{
					case TVN_SELCHANGED:
					{
						
						NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)lParam;
						pNMTreeView->itemNew.mask = TVIF_PARAM;
						if(SendDlgItemMessage(hwndDlg,IDC_TREE,TVM_GETITEM,TVGN_CARET,(LPARAM)&pNMTreeView->itemNew))
						{
							HRESULT hr;
							BSTR	nodeName;
							IXMLDOMElement* node;
							node = (IXMLDOMElement*)pNMTreeView->itemNew.lParam;

							hr = node->lpVtbl->get_nodeName(node,&nodeName);

							if(_tcsstr(nodeName, _T("book")) != NULL)
							{
								// warning: Best debugging tool ever:
								// MessageBeep(0);
								getBookData(hwndDlg, node);
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), TRUE);
								selectedCtl = NULL;
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), 0);
							}
							else if(_tcsstr(nodeName, _T("author")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_AUTHOR);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_AUTHOR);
							}
							else if(_tcsstr(nodeName, _T("title")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_TITLE);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_TITLE);
							}
							else if(_tcsstr(nodeName, _T("genre")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_GENRE);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_GENRE);
							}
							else if(_tcsstr(nodeName, _T("price")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_PRICE);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_PRICE);
							}
							else if(_tcsstr(nodeName, _T("publish_date")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_PUBLISHED);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_PUBLISHED);
							}
							else if(_tcsstr(nodeName, _T("description")) != NULL)
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								selectedCtl = GetDlgItem(hwndDlg,IDC_DESCRIPTION);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), IDC_DESCRIPTION);
							}
							else
							{
								EnableGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), FALSE);
								HiliteGroupboxControls(GetDlgItem(hwndDlg, IDC_PROPGROUP), 0);
							}
							SysFreeString(nodeName);
						}
						break;
					}
				}
			}
			break;
		}
        case WM_COMMAND:
            switch (GET_WM_COMMAND_ID(wParam, lParam))
            {
				case IDC_OPEN:
				{
					TCHAR *szFileName = malloc(sizeof(TCHAR) * MAX_PATH + 1);
					memset(szFileName, 0, sizeof(TCHAR) * MAX_PATH + 1);
					// Open XML file and populate tree
					if(GetOpenSaveDlg(szFileName, DLG_OPEN))
					{
						PopulateTree(GetDlgItem(hwndDlg, IDC_TREE), szFileName);
					}
					break;
				}
				case IDC_SAVE:
				{
					int _bufLen = GetWindowTextLength(GetDlgItem(hwndDlg,IDC_FILE));
					TCHAR * _szTextBuf = malloc(sizeof(TCHAR) * _bufLen + 1);
					GetDlgItemText(hwndDlg, IDC_FILE, _szTextBuf, _bufLen + 1);

					SaveTree(GetDlgItem(hwndDlg, IDC_TREE), _szTextBuf);
					free(_szTextBuf);
					break;
				}
				case IDC_SAVENODE:
				{
						HTREEITEM Selected = (HTREEITEM)SendMessage(GetDlgItem(hwndDlg, IDC_TREE),TVM_GETNEXTITEM,TVGN_CARET,(LPARAM)Selected);
						TV_ITEM tvI;
						tvI.mask = TVIF_PARAM;
						tvI.hItem = Selected;
						if(SendMessage(GetDlgItem(hwndDlg, IDC_TREE),TVM_GETITEM,0,(LPARAM)&tvI))
						{
							IXMLDOMElement* node;
							node = (IXMLDOMElement*)tvI.lParam;
							setBookData(hwndDlg, node);
						}
						break;
				}
				case IDC_COPY:
					CopyToClipboard(_T("Hello word!"));
					break;
                case IDC_CLOSE:
					ClearTree(GetDlgItem(hwndDlg, IDC_TREE));
                    EndDialog(hwndDlg, TRUE);
                    return TRUE;
            }
            break;
		case WM_CTLCOLOREDIT:
		{
			HWND hWnd = (HWND)lParam;
			if(hWnd != selectedCtl)
			{
				return FALSE;
			}
			HBRUSH hBrush = CreateSolidBrush(RGB(220,192,192));
			SetBkColor((HDC)wParam, RGB(220,192,192));
			SetTextColor((HDC)wParam, RGB(128,0,0));
			return (LRESULT)hBrush;
		}
		case WM_MOUSEMOVE:
		{
			OnMouseMove(GetDlgItem(hwndDlg, IDC_TREE), wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
			break;
		}
		case WM_LBUTTONDOWN:
		{
			OnLButtonDown(GetDlgItem(hwndDlg, IDC_TREE), wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
			break;
		}
		case WM_LBUTTONUP:
		{
			OnLButtonUp(GetDlgItem(hwndDlg, IDC_TREE), wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
			break;
		}
        case WM_CLOSE:
			ClearTree(GetDlgItem(hwndDlg, IDC_TREE));
            EndDialog(hwndDlg, 0);
            return TRUE;
    }

    return FALSE;
}

///////////////////////////////////////////////////////////////////////////////////
//                                                                               //
// This function sets the data in some of the controls of the dialog on to the   //
// XML                                                                           //
//                                                                               //
///////////////////////////////////////////////////////////////////////////////////
static BOOL setBookData(HWND hwndDlg, IXMLDOMElement* node)
{
	HRESULT hr = S_OK;

	//////////////////////////////////////////////////////////////////////
	// selectSingleNode method is direct and order-less better than     //
	// lookup method but does not adhere to the W3C Recommendation      //
	// Can also use XPath Navigation to look deeper                     //
	//////////////////////////////////////////////////////////////////////
	
	// Author field
	IXMLDOMNode *_authorNode;
	hr = node->lpVtbl->selectSingleNode(node, L"author", &_authorNode);
	if(SUCCEEDED(hr))
	{
		HWND _hwndEditBox = GetDlgItem(hwndDlg,IDC_AUTHOR);
		int _bufLen = GetWindowTextLength(_hwndEditBox);
		TCHAR * _szTextBuf = malloc(sizeof(TCHAR) * _bufLen + 1);
		GetDlgItemText(hwndDlg, IDC_AUTHOR, _szTextBuf, _bufLen + 1);
		_authorNode->lpVtbl->put_text(_authorNode, _szTextBuf);
		free(_szTextBuf);
		_authorNode->lpVtbl->Release(_authorNode);
	}
	// Title field
	IXMLDOMNode *_titleNode;                                    
	hr = node->lpVtbl->selectSingleNode(node, L"title", &_titleNode);
	if(SUCCEEDED(hr))
	{
		HWND _hwndEditBox = GetDlgItem(hwndDlg,IDC_TITLE);
		int _bufLen = GetWindowTextLength(_hwndEditBox);
		TCHAR * _szTextBuf = malloc(sizeof(TCHAR) * _bufLen + 1);
		GetDlgItemText(hwndDlg, IDC_TITLE, _szTextBuf, _bufLen + 1);
		_titleNode->lpVtbl->put_text(_titleNode, _szTextBuf);
		free(_szTextBuf);
		_titleNode->lpVtbl->Release(_titleNode);
	}
	// Genere field
	IXMLDOMNode *_genreNode;                                    
	hr = node->lpVtbl->selectSingleNode(node, L"genre", &_genreNode);
	if(SUCCEEDED(hr))
	{
		int iSel = SendMessage(GetDlgItem(hwndDlg,IDC_GENRE),CB_GETCURSEL,0,0);
		_genreNode->lpVtbl->put_text(_genreNode, comboGenres[iSel]);
		_genreNode->lpVtbl->Release(_genreNode);
	}
	return SUCCEEDED(hr);
}
///////////////////////////////////////////////////////////////////////////////////
//                                                                               //
// This function gets data from XML nodes and pupulates the dialog controls      //
// It takes a IXMLDOMElement and looks for the right node to put data on screen  //
// We use several methods to retreive the value or text form a specific node     //
// Read each example carefully                                                   //
//                                                                               //
///////////////////////////////////////////////////////////////////////////////////
static BOOL getBookData(HWND hwndDlg, IXMLDOMElement* node)
{
	HRESULT 		hr = S_OK;
	IXMLDOMNodeList		*pXMLNodeList;
	IXMLDOMNode* nodeItem = NULL;

	//////////////////////////////////////////////////////////////////////
	// selectSingleNode method is direct and order-less better than     //
	// lookup method but does not adhere to the W3C Recommendation      //
	// Can also use XPath Navigation to look deeper                     //
	//////////////////////////////////////////////////////////////////////
	// Author field                                                     //
	IXMLDOMNode *authorNode;                                            //
	hr = node->lpVtbl->selectSingleNode(node, L"author", &authorNode);  //
	if(SUCCEEDED(hr))                                                   //
	{                                                                   //
		BSTR author;                                                    //
		authorNode->lpVtbl->get_text(authorNode,&author);               //
		authorNode->lpVtbl->Release(authorNode);                        //
#ifndef _UNICODE                                                        //
		char *szBufauthor = BSTRtoANSI(author);                         //
		SetDlgItemText(hwndDlg, IDC_AUTHOR, szBufauthor);               //
		free(szBufauthor);                                              //
#else                                                                   //
		SetDlgItemText(hwndDlg, IDC_AUTHOR, author);                    //
#endif                                                                  //
		SysFreeString(author);                                          //
	}                                                                   //
	//////////////////////////////////////////////////////////////////////


	//////////////////////////////////////////////////////////////////////
	// Work with indexed access of IXMLDOMNodeList interface            // 
	// XML file must follow an exact order                              //
	//////////////////////////////////////////////////////////////////////
	node->lpVtbl->get_childNodes(node, &pXMLNodeList);

	// Title field
	BSTR title;
	pXMLNodeList->lpVtbl->get_item(pXMLNodeList,1, &nodeItem);
	nodeItem->lpVtbl->get_text(nodeItem,&title);
	nodeItem->lpVtbl->Release(nodeItem);
#ifndef _UNICODE
	char *szBuftitle = BSTRtoANSI(title);
	SetDlgItemText(hwndDlg, IDC_TITLE, szBuftitle);
	free(szBuftitle);
#else
	SetDlgItemText(hwndDlg, IDC_TITLE, title);
#endif
	SysFreeString(title);

	// Genere field
	BSTR genre;
	pXMLNodeList->lpVtbl->get_item(pXMLNodeList,2, &nodeItem);
	nodeItem->lpVtbl->get_text(nodeItem,&genre);
	nodeItem->lpVtbl->Release(nodeItem);
	
	TCHAR *_szGenre;
#ifndef _UNICODE
	_szGenre = BSTRtoANSI(genre);
#else
	_szGenre = genre;
#endif
	for(int i = 0; i < MAX_GENRES; i++)
	{
		if(_tcsstr(_szGenre, comboGenres[i]) != NULL)
		{
			SendDlgItemMessage(hwndDlg, IDC_GENRE, CB_SETCURSEL, i, 0);
			break;
		}
	}
#ifndef _UNICODE
	free(_szGenre);
#endif
	SysFreeString(genre);

	// Price field
	BSTR price;
	pXMLNodeList->lpVtbl->get_item(pXMLNodeList,3, &nodeItem);
	nodeItem->lpVtbl->get_text(nodeItem,&price);
	nodeItem->lpVtbl->Release(nodeItem);
#ifndef _UNICODE
	char *szBufprice = BSTRtoANSI(publish_date);
	SetDlgItemText(hwndDlg, IDC_PRICE, szBufprice);
	free(szBufprice);
#else
	SetDlgItemText(hwndDlg, IDC_PRICE, price);
#endif
	SysFreeString(price);

	//////////////////////////////////////////////////////////////////////
	// Work time format & convertion to SYSTEMTIME                      //
	//////////////////////////////////////////////////////////////////////
	BSTR publish_date;
	pXMLNodeList->lpVtbl->get_item(pXMLNodeList,4, &nodeItem);
	nodeItem->lpVtbl->get_text(nodeItem,&publish_date);
	nodeItem->lpVtbl->Release(nodeItem);
	
	SYSTEMTIME stime;
	memset(&stime,0,sizeof(SYSTEMTIME));
	SendDlgItemMessage(hwndDlg, IDC_PUBLISHED, DTM_GETSYSTEMTIME, 0, (LPARAM) &stime);

	//////////////////////////////////////////////////////////
	// _stscanf -> sscanf if ANSI or swscanf if UNICODE     //
	// The following code could be done in a simpler        //
	// way or maybe more appeling to the eye but I          //
	// decided to keep it this way.                         //
	//////////////////////////////////////////////////////////
#ifndef _UNICODE
	char *szBufStime = BSTRtoANSI(publish_date);
	sscanf(szBufStime,"%d-%d-%d-%d.%d.%d",
               &stime.wYear,
               &stime.wMonth,
               &stime.wDay,
               &stime.wHour,
               &stime.wMinute,
               &stime.wSecond);
   free(szBufStime);
#else
	_stscanf(publish_date,_T("%d-%d-%d-%d.%d.%d"),
               &stime.wYear,
               &stime.wMonth,
               &stime.wDay,
               &stime.wHour,
               &stime.wMinute,
               &stime.wSecond);
#endif
	SendDlgItemMessage(hwndDlg, IDC_PUBLISHED, DTM_SETSYSTEMTIME, GDT_VALID, (LPARAM) &stime);
	SysFreeString(publish_date);

	// Description field
	BSTR description;
	pXMLNodeList->lpVtbl->get_item(pXMLNodeList,5, &nodeItem);
	nodeItem->lpVtbl->get_text(nodeItem,&description);
	//nodeItem->lpVtbl->put_text(nodeItem, L"holaaaa");
	nodeItem->lpVtbl->Release(nodeItem);
#ifndef _UNICODE
	char *szBufdescription = BSTRtoANSI(description);
	SetDlgItemText(hwndDlg, IDC_DESCRIPTION, szBufdescription);
	free(szBufdescription);
#else
	SetDlgItemText(hwndDlg, IDC_DESCRIPTION, description);
#endif
	SysFreeString(description);

	// We are not working with this object any more, release it
	pXMLNodeList->lpVtbl->Release(pXMLNodeList);

	//////////////////////////////////////////////////////////////////////
	// Work with attribute nodes IXMLDOMNamedNodeMap Interface          //
	//////////////////////////////////////////////////////////////////////
	IXMLDOMNamedNodeMap* namedNodeMap = NULL;
	IXMLDOMNode*	ppNamedItem = NULL;
	hr = node->lpVtbl->get_attributes(node,&namedNodeMap);
	
	hr = namedNodeMap->lpVtbl->getNamedItem(namedNodeMap,L"id",&ppNamedItem);
	if(!ppNamedItem)
	{
		namedNodeMap->lpVtbl->Release(namedNodeMap);
		return FALSE; // in this case is allways FALSE
	}
	BSTR isbn;
	ppNamedItem->lpVtbl->get_text(ppNamedItem,&isbn);
#ifndef _UNICODE
	char *szBufisbn = BSTRtoANSI(isbn);
	SetDlgItemText(hwndDlg, IDC_ISBN, szBufisbn);
	free(szBufisbn);
#else
	SetDlgItemText(hwndDlg, IDC_ISBN, isbn);
#endif
	SysFreeString(isbn);
	// We have to release all references
	// It's like this all the time, chances are there is going to be a leak somewhere
	ppNamedItem->lpVtbl->Release(ppNamedItem);
	namedNodeMap->lpVtbl->Release(namedNodeMap);

	return TRUE;
}

int EnableGroupboxControls(HWND hWnd, BOOL bEnable)
{
    int rc = 0;

    if (IsWindow(hWnd))
    {
        // get class name
        TCHAR szClassName[MAX_PATH];
        szClassName[0] = _T('\0');
        GetClassName(hWnd, szClassName, sizeof(szClassName)/sizeof(TCHAR)-2);

        // get window style
        LONG lStyle = GetWindowLong(hWnd, GWL_STYLE);

        if ((_tcsicmp(szClassName, _T("Button")) == 0) &&
            ((lStyle & BS_TYPEMASK) == BS_GROUPBOX))
        {
            // this is a groupbox

            RECT rectGroupbox;
            GetWindowRect(hWnd, &rectGroupbox);

            // get first child control

            HWND hWndChild = 0;
            HWND hWndParent = GetParent(hWnd);
            if (IsWindow(hWndParent))
                hWndChild = GetWindow(hWndParent, GW_CHILD);

            while (hWndChild)
            {
                RECT rectChild;
                GetWindowRect(hWndChild, &rectChild);

                // check if child rect is entirely contained within groupbox
                if ((rectChild.left >= rectGroupbox.left) &&
                    (rectChild.right <= rectGroupbox.right) &&
                    (rectChild.top >= rectGroupbox.top) &&
                    (rectChild.bottom <= rectGroupbox.bottom))
                {
                    //TRACE(_T("found child window 0x%X\n"), hWndChild);
                    EnableWindow(hWndChild, bEnable);
                    rc++;
                }

                // get next child control
                hWndChild = GetWindow(hWndChild, GW_HWNDNEXT);
            }

            // if any controls were affected, invalidate the parent rect
            if (rc && IsWindow(hWndParent))
            {
                InvalidateRect(hWndParent, NULL, FALSE);
            }
        }
    }
    return rc;
}

int HiliteGroupboxControls(HWND hWnd, int id)
{
    int rc = 0;

    if (IsWindow(hWnd))
    {
        // get class name
        TCHAR szClassName[MAX_PATH];
        szClassName[0] = _T('\0');
        GetClassName(hWnd, szClassName, sizeof(szClassName)/sizeof(TCHAR)-2);

        // get window style
        LONG lStyle = GetWindowLong(hWnd, GWL_STYLE);

        if ((_tcsicmp(szClassName, _T("Button")) == 0) &&
            ((lStyle & BS_TYPEMASK) == BS_GROUPBOX))
        {
            // this is a groupbox

            RECT rectGroupbox;
            GetWindowRect(hWnd, &rectGroupbox);

            // get first child control

            HWND hWndChild = 0;
            HWND hWndParent = GetParent(hWnd);
            if (IsWindow(hWndParent))
                hWndChild = GetWindow(hWndParent, GW_CHILD);

            while (hWndChild)
            {
                RECT rectChild;
                GetWindowRect(hWndChild, &rectChild);

                // check if child rect is entirely contained within groupbox
                if ((rectChild.left >= rectGroupbox.left) &&
                    (rectChild.right <= rectGroupbox.right) &&
                    (rectChild.top >= rectGroupbox.top) &&
                    (rectChild.bottom <= rectGroupbox.bottom))
                {
                    if(id == GetDlgCtrlID(hWndChild))
					{
						EnableWindow(hWndChild, TRUE);
					}
					else
					{
						RedrawWindow(hWndChild, NULL, NULL, RDW_INVALIDATE|RDW_FRAME);
					}
                    //EnableWindow(hWndChild, bEnable);
                    rc++;
                }

                // get next child control
                hWndChild = GetWindow(hWndChild, GW_HWNDNEXT);
            }

            // if any controls were affected, invalidate the parent rect
            if (rc && IsWindow(hWndParent))
            {
                InvalidateRect(hWndParent, NULL, FALSE);
            }
        }
    }
    return rc;
}

void CopyToClipboard(TCHAR *buf)
{
	if(OpenClipboard(0))
	{
	// Empty what's in there...
	EmptyClipboard();
	// Allocate global memory for transfer...
	HGLOBAL hText = GlobalAlloc(GMEM_MOVEABLE|GMEM_DDESHARE, sizeof(TCHAR) * _tcslen(buf)+4);
	// Put your string in the global memory...
	TCHAR *ptr = (TCHAR *)GlobalLock(hText);
	_tcscpy(ptr, buf);
	GlobalUnlock(hText);
        
	SetClipboardData(CF_UNICODETEXT, hText);
        
	CloseClipboard();
	// Free memory...
	GlobalFree(hText);
	}     
}
///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Open the file open/save dialog and get the selected file       //
// Input:         buffer holds the file path must be MAX_PATH size               //
//                buflen > MAX_PATH                                              //
//                option (0, open file, 1, save file)                            //
//                                                                               //
// Output:        LPCTSTR file full path                                         //
// Errors:        If the function succeeds, the return value is buffer           //
//                If the function fails, the return value is 0.                  //
// Author:        Benjamin Maggi 2006 / benjaminmaggi@gmail.com                  //
//                Converted to UNICODE compatible 2009                           //
///////////////////////////////////////////////////////////////////////////////////
BOOL GetOpenSaveDlg(TCHAR *buffer, BYTE option)
{
	TCHAR tmpfilter[39];
	int i = 0;
	OPENFILENAME ofn;

	memset(&ofn,0,sizeof(ofn));
	
	ofn.lStructSize		= sizeof(ofn);
	ofn.hInstance		= GetModuleHandle(NULL);
	ofn.hwndOwner		= GetActiveWindow();
	ofn.lpstrFile		= buffer;
	ofn.nMaxFile		= MAX_PATH;
	ofn.nFilterIndex	= 2;
	ofn.lpstrDefExt		= _T("xml");
	
	//////////////////////////////////////////////////////
	// File type filters                                //
	//////////////////////////////////////////////////////
	_tcscpy(tmpfilter, _T("All files,*.*,XML files,*.xml"));

	while(tmpfilter[i]) 
	{
		if (tmpfilter[i] == _T(','))
			tmpfilter[i] = 0;
		i++;
	}
	tmpfilter[i++]	= 0; tmpfilter[i++] = 0;
	ofn.Flags		= 539678;
	ofn.lpstrFilter = tmpfilter;


	if(option == 0)
	{
		ofn.lpstrTitle = _T("Open XML");
		if(GetOpenFileName(&ofn))
		{
			return TRUE;
		}
	}
	else
	{
		ofn.lpstrTitle = _T("Save XML");
		if(GetSaveFileName(&ofn))
		{
			return TRUE;
		}
	}
	return FALSE;
}
