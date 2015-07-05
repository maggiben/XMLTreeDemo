///////////////////////////////////////////////////////////////////////////////////
//                                                                               //
//                       THIS IS THE RELEASE VERSION                             //
//                                                                               //
///////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////
// File Name:           XMLTree.c                                                //
// Author:              Benjamin Maggi                                           //
// Descripcion:         Read XML using MSXML and pupulates a tree control        //
// Org. Date:           13/03/2008                                               //
// Last Modification:   22/05/2009                                               //
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
// Rev.1                                                                         //
//      * Mostly Unicode compatible                                              //
// Rev.2                                                                         //
//      * Save tree data to XML                                                  //
// Rev.3                                                                         //
//      * Better interface added: _AddComment & _AddPI functions                 //
//                                                                               //
// TO-DO:                                                                        //
//      * Inset/move/edit items in a tree and then build/modify a node form it   //
//      * Support Drag&Drop primary to allow arragement of items                 //
///////////////////////////////////////////////////////////////////////////////////

#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <oleauto.h>					// Ole automation
#include <stdio.h>
#include <tchar.h>                      // Character set independent macros for text manipulation
#include <wchar.h>                      // Wide Char
#include <msxml.h>                      // MSXML component header
#include "XMLTree.h"
#include "main.h"

///////////////////////////////////////////////////////////////////////////////////
// Nessasary libs                                                                //
///////////////////////////////////////////////////////////////////////////////////

#pragma comment(lib, "kernel32.lib")	// Memory manipulation
#pragma comment(lib, "delayimp.lib")	// Delay Loading a DLL
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")		// GDI Manipulation
#pragma comment(lib, "comctl32.lib")	// Windows Connmon controls
#pragma comment(lib, "comdlg32.lib")	// Windows Connmon dialogs
#pragma comment(lib, "advapi32.lib")	// Windows Registry access
#pragma comment(lib, "uuid.lib")		// Contains the GUID for various classes 
#pragma comment(lib, "ole32.lib")		// 
#pragma comment(lib, "oleaut32.lib")	// OLE Automation 
#pragma comment(lib, "shell32.lib")		// Window Shell helper funcitons
#pragma comment(lib, "oldnames.lib")	
#pragma comment(lib, "urlmon.lib")	

///////////////////////////////////////////////////////////////////////////////////
// Internal Variables                                                            //
///////////////////////////////////////////////////////////////////////////////////
// Once we load an XML these two variables should remain valid
static IXMLDOMDocument* _xmlDocument = NULL;
static IXMLDOMElement*  _xmlRoot = NULL;
// Drag and Drop
static BOOL _bDragging = FALSE;
static HTREEITEM _HitTarget;
static HTREEITEM _Selected;
///////////////////////////////////////////////////////////////////////////////////
// Internal Functions                                                            //
///////////////////////////////////////////////////////////////////////////////////
// TreeView Specific Functions
static BOOL                  _DeleteTreeItem      (HWND hWndTree);
static BOOL                  _AddSiblingTreeItem  (HWND hWndTree, IXMLDOMDocument *xmlDoc);
static BOOL                  _AddChildTreeItem    (HWND hWndTree, IXMLDOMDocument *xmlDoc);
static BOOL 				 _PopulateNode        (IXMLDOMElement *node, HTREEITEM hParent, HTREEITEM hInsAfter, HWND tree);
static BOOL 				 _PopulateAttributes  (IXMLDOMElement *node, HTREEITEM hParent, HTREEITEM hInsAfter, HWND tree);
// XML Specific Functions
static IXMLDOMDocument      *_GetDomDocument      (TCHAR *fileName);
static IXMLDOMElement       *_GetRootElement      (IXMLDOMDocument* pXMLDoc);
static IXMLDOMNodeList      *_GetChildNodes       (IXMLDOMElement* domElement);
static IXMLDOMElement       *_AddRootElement      (IXMLDOMDocument *xmlDoc);
static BOOL                  _AddChild            (IXMLDOMDocument *xmlDoc, IXMLDOMNode *child);
static BOOL                  _AddComment          (IXMLDOMDocument *xmlDoc, TCHAR *comment);
static BOOL                  _AddPI               (IXMLDOMDocument *xmlDoc);
static BOOL                  _CopyNode            (IXMLDOMNode *pXmlDest, IXMLDOMNode *pXmlSrc);
static IXMLDOMDocument      *_CreateDocument      (void);

//
//  Deletes the currently selectd node.
//
static BOOL _DeleteTreeItem(HWND hWndTree) 
{
    HTREEITEM hCurrent = TreeView_GetSelection(hWndTree);
    if (!hCurrent)
        return FALSE;
    
    return TreeView_DeleteItem(hWndTree, hCurrent);
}
//
//  Adds a sibling to a selected node.
//
static BOOL _AddSiblingTreeItem(HWND hWndTree, IXMLDOMDocument *xmlDoc)
{
    HTREEITEM hCurrent = TreeView_GetSelection(hWndTree);
    HTREEITEM hParent  = TreeView_GetParent(hWndTree, hCurrent);

    TVINSERTSTRUCT tvs;
    tvs.item.mask                   = TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT;
    if (!hCurrent || !hParent)
    {
        tvs.hParent = TVI_ROOT;
        tvs.item.pszText            = TEXT("Root Item");
    }
    else
    {
        tvs.hParent = hParent;
        tvs.item.pszText            = TEXT("Sibling Item");
    }

    tvs.item.cchTextMax             = lstrlen(tvs.item.pszText) + 1;
    tvs.hInsertAfter = TVI_LAST;
    //tvs.item.iImage                 = rand() % ILCOUNT;
    //tvs.item.iSelectedImage         = rand() % ILCOUNT;

    TreeView_InsertItem(hWndTree, &tvs);
    
    if (hCurrent)
        return TreeView_Expand(hWndTree, hCurrent, TVE_EXPAND);    
}

//
//  Adds one child to selected node or if nothing is selected
//  adds another root node.
//
static BOOL _AddChildTreeItem(HWND hWndTree, IXMLDOMDocument *xmlDoc)
{
    HTREEITEM hCurrent = TreeView_GetSelection(hWndTree);
    
    TVINSERTSTRUCT tvs;
    tvs.item.mask                   = TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT;
    if (!hCurrent)
    {
        tvs.hParent = TVI_ROOT;
        tvs.item.pszText            = TEXT("Root Item");    
    }
    else
    {
        tvs.hParent = hCurrent;
        tvs.item.pszText            = TEXT("Child Item");
    }

    tvs.item.cchTextMax             = lstrlen(tvs.item.pszText) + 1;
    tvs.hInsertAfter = TVI_LAST;
  //  tvs.item.iImage                 = rand() % ILCOUNT;
  //  tvs.item.iSelectedImage         = rand() % ILCOUNT;


    HTREEITEM hNewItem = TreeView_InsertItem(hWndTree, &tvs);
    
    if (hCurrent)
        TreeView_Expand(hWndTree, hCurrent, TVE_EXPAND);

    if (hNewItem)
        TreeView_Select(hWndTree, hNewItem, TVGN_CARET);

}

///////////////////////////////////////////////////////////////////////////////////
// Set some images to add visuals                                                //
///////////////////////////////////////////////////////////////////////////////////
static BOOL                  _SetTreeImages       (HWND treeHwd, HINSTANCE hInst);
///////////////////////////////////////////////////////////////////////////////////
// Ole string helpers                                                            //
// You may decide to put this outside this file for reusability or use your own  //
// conversion functions                                                          //
///////////////////////////////////////////////////////////////////////////////////
static BSTR                   ANSItoBSTR          (char *ansistr);
static char                  *BSTRtoANSI          (BSTR bstr);
static VARIANT                VariantString       (BSTR str);
///////////////////////////////////////////////////////////////////////////////////
// Error handling and end user messages                                          //
///////////////////////////////////////////////////////////////////////////////////
static TCHAR *               _ErrorString        (DWORD errno);


int OnMouseMove(HWND hWnd, WPARAM wParam, int xPos, int yPos)
{
	POINT pnt ;
	HTREEITEM h_item = NULL ;

	pnt.x = xPos;//GET_X_LPARAM(l_parm) ;
	pnt.y = yPos;//GET_Y_LPARAM(l_parm) ;
	if(_bDragging)
	{
      //unlock window and allow updates to occur
      ImageList_DragLeave(NULL) ;
      ClientToScreen(GetParent(hWnd), &pnt) ;
      //check with the tree control to see if we are on an item
      TVHITTESTINFO tv_ht ;
      ZeroMemory(&tv_ht, sizeof(TVHITTESTINFO)) ;
      tv_ht.flags = TVHT_ONITEM ;
      tv_ht.pt.x = pnt.x ;
      tv_ht.pt.y = pnt.y ;
      ScreenToClient(hWnd, &(tv_ht.pt)) ;
      h_item = (HTREEITEM)SendMessage(hWnd, TVM_HITTEST, 0, (LPARAM)&tv_ht) ;
      if(h_item)
      {  //if we had a hit, then drop highlite the item
         SendMessage(hWnd, TVM_SELECTITEM, TVGN_DROPHILITE, (LPARAM)h_item) ;
      }
      //paint the image in the new location
      ImageList_DragMove(pnt.x,pnt.y) ;
      //lock the screen again
      ImageList_DragEnter(NULL, pnt.x, pnt.y) ;
	}
	return 0;
}
int	OnLButtonDown(HWND hWnd, WPARAM wParam, int xPos, int yPos)
{
	ReleaseCapture();
	SendMessage(hWnd,WM_NCLBUTTONDOWN,HTCAPTION,0);
	return 0;
}
int	OnLButtonUp(HWND hWnd, WPARAM wParam, int xPos, int yPos)
{
	if (_bDragging) 
	{
		ImageList_DragLeave(hWnd);
		ImageList_EndDrag();
		_Selected = (HTREEITEM)SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_DROPHILITE, 0);
		SendMessage(hWnd, TVM_SELECTITEM, TVGN_CARET, (LPARAM)_Selected);
		SendMessage(hWnd, TVM_SELECTITEM, TVGN_DROPHILITE, 0);
		ReleaseCapture();
		ShowCursor(TRUE); 
		_bDragging = FALSE;
	}
	return 0;
}
///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Reflection from parent WM_NOTIFY message                       //
// Input:         Messages reflected from main window proc                       //
//                                                                               //
// Notes:         We need to catch 2 important messages before our application   //
//                TVN_ITEMEXPANDING and TVN_DELETEITEM to-do catch insert        //
//                and append nodes to do it more realistically                   //
///////////////////////////////////////////////////////////////////////////////////
int OnNotify(HWND hWnd, UINT message, int idCtrl, NMHDR *pnmh)
{
	switch(pnmh->code)
	{
		case TVN_ITEMEXPANDING:
		{
			//////////////////////////////////////////////////////
			// Implement On-Demand node load                    //
			// lParam stores the node get it                    //
			//////////////////////////////////////////////////////
			HRESULT hr;
			NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pnmh;
			HTREEITEM hItem = pNMTreeView->itemNew.hItem;
			IXMLDOMElement* node;
			pNMTreeView->itemNew.mask = TVIF_PARAM;
			HWND treeHwnd = pnmh->hwndFrom;
			if(SendMessage(treeHwnd, TVM_GETITEM, 0, (LPARAM)&pNMTreeView->itemNew))
			{
				// Catch specific nodes here may also put some call back to 
				// For example restrict users to certain deep level
				// I wanted to implement such kind if feature
				// But really didn't had time, you can start from here:
				BSTR nodeName;
				node = (IXMLDOMElement*)pNMTreeView->itemNew.lParam;
				hr = node->lpVtbl->get_nodeName(node,&nodeName);
				SysFreeString(nodeName);
			}
			else
			{
				return TRUE;
			}
			LockWindowUpdate(hWnd);
			if (pNMTreeView->action == TVE_EXPAND) 
			{
				// Delete first empty item
				HTREEITEM hChildItem;
				if ((hChildItem = TreeView_GetChild(treeHwnd,hItem)) != NULL) 
				{
					TreeView_DeleteItem(treeHwnd, hChildItem);
				}

				VARIANT_BOOL hasChild;
				hr = node->lpVtbl->hasChildNodes(node, &hasChild);
				if (SUCCEEDED(hr) && hasChild)
				{
					_PopulateNode((IXMLDOMElement *)node, hItem,0,treeHwnd);
				}
				if(FAILED(hr))
				{
					MessageBoxA(NULL,__func__, "Error get_attributes",MB_OK|MB_ICONERROR);
				}

			}
			else if(pNMTreeView->action == TVE_COLLAPSE)
			{
				HTREEITEM hChildItem;
				if((hChildItem = TreeView_GetChild(treeHwnd,hItem)) == NULL) 
				{
					break;
				}
				do {
						HTREEITEM hNextItem = TreeView_GetNextSibling(treeHwnd,hChildItem);
						TV_ITEM tvI;
						tvI.mask = TVIF_PARAM;
						tvI.hItem = hChildItem;
						SendMessage(treeHwnd, TVM_GETITEM, 0, (LPARAM)&tvI);
						// TVN_DELETEITEM will take care of object release
						// IXMLDOMElement * node;
						// node = (IXMLDOMElement *)tvI.lParam;
						// node->lpVtbl->Release(node);
						
						// Finally delete the item
						TreeView_DeleteItem(treeHwnd,hChildItem);
						hChildItem = hNextItem;
					} while (hChildItem != NULL);

				//Insert an empty element 
				hChildItem = InsertTreeItem(hItem,NULL,0,0,treeHwnd,NULL);
			}
			LockWindowUpdate(NULL);
			break;
		}
		case TVN_DELETEITEM:
		{
			HRESULT hr;
			HWND treeHwnd = pnmh->hwndFrom;
			NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pnmh;
			HTREEITEM hChildItem;
			if((hChildItem = TreeView_GetChild(treeHwnd,pNMTreeView->itemOld.hItem)) == NULL) 
			{
				break;
			}

			do
			{
				HTREEITEM hNextItem = TreeView_GetNextSibling(treeHwnd,hChildItem);
				//////////////////////////////////////////////////////
				// Release the XML node we no longer need it        //
				// This will prevent leaking because we dont reuse  //
				// The objects rather than create new ones          //
				//////////////////////////////////////////////////////
				IXMLDOMElement * node;
				TV_ITEM tvI;
				tvI.mask = TVIF_PARAM;
				tvI.hItem = hChildItem;
				TreeView_GetItem(treeHwnd,&tvI);
				node = (IXMLDOMElement *)tvI.lParam;
				hr = node->lpVtbl->Release(node);
				//////////////////////////////////////////////////////
				// Finally delete the item                          //
				//////////////////////////////////////////////////////
				TreeView_DeleteItem(treeHwnd,hChildItem);
				hChildItem = hNextItem;
			}
			while (hChildItem != NULL);
			IXMLDOMElement* node;
			if(pNMTreeView->itemOld.lParam)
			{
				node = (IXMLDOMElement*)pNMTreeView->itemOld.lParam;
				node->lpVtbl->Release(node);
			}
		}
		case TVN_BEGINDRAG:
		{
			/*
			HIMAGELIST hImg;
			LPNMTREEVIEW lpnmtv = (LPNMTREEVIEW)pnmh;
			hImg = TreeView_CreateDragImage(pnmh->hwndFrom, lpnmtv->itemNew.hItem);
			ImageList_BeginDrag(hImg, 0, 0, 0);
			ImageList_DragEnter(pnmh->hwndFrom, lpnmtv->ptDrag.x, lpnmtv->ptDrag.y);
			//ShowCursor(FALSE); 
			SetCapture(hWnd); 
			_bDragging = TRUE;
			*/
			LRESULT pResult;
			OnBegindragTree(pnmh, &pResult);
			return pResult;
		}
	}
	return TRUE;
}

// Inspired in Paul S. Vickery C++ sources
void OnBegindragTree(NMHDR* pnmh, LRESULT* pResult)
{
	LPNMTREEVIEW lpnmtv = (LPNMTREEVIEW)pnmh;
	HIMAGELIST hImg;
	//Create an image list that holds our drag image
	hImg = TreeView_CreateDragImage(pnmh->hwndFrom, lpnmtv->itemNew.hItem) ;
	//begin the drag operation
	ImageList_BeginDrag(hImg, 0, 0, 0) ;
	//hide the cursor
	ShowCursor(FALSE) ;
	//capture the mouse
	SetCapture(GetParent(pnmh->hwndFrom)) ;
	//convert coordinates to screen coordinates
	ClientToScreen(pnmh->hwndFrom, &(lpnmtv->ptDrag)) ;
	//paint our drag image and lock the screen.
	ImageList_DragEnter(NULL, lpnmtv->ptDrag.x, lpnmtv->ptDrag.y) ;
	UpdateWindow(pnmh->hwndFrom);
	_bDragging = TRUE;
	
	return;
}
///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Fills the SysTreeView with data from XML file                  //
// Input:         HWND of the parent window                                      //
// Input:         TCHAR *fileName: Name of the file to load                      //
//                                                                               //
// Output:        Pupulates the tree with first child in the node list           //
//                The rest of the items will be inserted on demand               //
// Errors:        If the function succeeds, the return value is TRUE             //
//                If the function fails, the return value is FALSE.              //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
BOOL PopulateTree(HWND tree, TCHAR *fileName)
{
	HRESULT hr;
	BSTR nodeName;
	TCHAR *errMsg;

	if(_xmlDocument || _xmlRoot)
	{
		if(MessageBox(0,_T("Another XML file is open\rAny unsaved data will be lost"),_T("Error PopulateTree() !"),MB_YESNO | MB_ICONSTOP | MB_SETFOREGROUND) == IDYES)
		{
			ClearTree(tree);
		}
		else
		{
			return FALSE;
		}
	}
	_SetTreeImages(tree, GetModuleHandle(NULL));

	CoInitialize(NULL);
	_xmlDocument = _GetDomDocument(fileName);
	if(_xmlDocument == NULL)
	{
		return FALSE;
	}
	_xmlRoot = _GetRootElement(_xmlDocument);
	//////////////////////////////////////////////////////
	// Get the name of the root node                    //
	//////////////////////////////////////////////////////
	hr = _xmlRoot->lpVtbl->get_nodeName(_xmlRoot,&nodeName);
	if(FAILED(hr)) { goto clean; }
#ifndef _UNICODE
	char *	tmpStr;
	tmpStr = BSTRtoANSI(nodeName);
	InsertTreeItem(0,tmpStr,0,0,tree,_xmlRoot);
	free(tmpStr);
#else
	InsertTreeItem(0,nodeName,0,0,tree,_xmlRoot);
#endif
	SysFreeString(nodeName);

	return TRUE;
clean:
 	errMsg = _ErrorString(hr);
	MessageBox(0,errMsg,_T("Error in CoCreateInstance() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
	LocalFree(errMsg);

	if (_xmlRoot) _xmlRoot->lpVtbl->Release(_xmlRoot);
	if (_xmlDocument) _xmlDocument->lpVtbl->Release(_xmlDocument);
	CoUninitialize();
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////////
// Release _xmlRoot & _xmlDocument also call CoUninitialize()                    //
// You would use this call in WM_NCDESTROY or before exiting the thread          //
///////////////////////////////////////////////////////////////////////////////////
BOOL ClearTree(HWND hwndTree)
{
	HRESULT hr;

	// Clear the tree's image list first
	HIMAGELIST 	hImageList = NULL;
	hImageList = (HIMAGELIST)SendMessage(hwndTree, TVM_GETIMAGELIST, TVSIL_NORMAL, (LPARAM)0);
	if(hImageList)
	{
		ImageList_Destroy(hImageList);
	}

	if (_xmlRoot && _xmlDocument)
	{
		int TreeCount = TreeView_GetCount(hwndTree);
		for(int i = 0; i < TreeCount; i++)
		{
			TreeView_DeleteAllItems(hwndTree);
		}
		hr = _xmlRoot->lpVtbl->Release(_xmlRoot);
		hr = _xmlDocument->lpVtbl->Release(_xmlDocument);

		// Set our global vaiables to NULL
		_xmlRoot = NULL;
		_xmlDocument = NULL;
	}
	CoUninitialize();
	return SUCCEEDED(hr);
}
///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Saves tree data to XML.                                        //
// Input:         HWND hWndTree: handle of the Tree Control                      //
// Input:         TCHAR *filename: Name of the file to save                      //
//                                                                               //
//                                                                               //
// Output:        A new XML file or rewrite in case of existing                  //                 //
// Errors:        If the function succeeds, the return value is TRUE             //
//                If the function fails, the return value is FALSE.              //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
BOOL SaveTree(HWND hWndTree, TCHAR *filename)
{
	HRESULT	hr;
	IXMLDOMDocument* XMLDoc = NULL;
	CoInitialize(NULL);

	XMLDoc = _CreateDocument();

	// Get the tree root item
	HTREEITEM rootItem = TreeView_GetRoot(hWndTree);
	TV_ITEM tvI;
	tvI.mask = TVIF_PARAM;
	tvI.hItem = rootItem;
	if(SendMessage(hWndTree,TVM_GETITEM,0,(LPARAM)&tvI))
	{
		MessageBeep(0);

		IXMLDOMElement* node;
		node = (IXMLDOMElement*)tvI.lParam;

		_AddPI(XMLDoc);
		_AddComment(XMLDoc, _T("This file was written with by SaveTree2()"));
		hr = XMLDoc->lpVtbl->appendChild(XMLDoc, (void *)node, NULL);
	}
	else
	{
		MessageBoxA(NULL,__func__, "Error No Item Selected!",MB_OK|MB_ICONERROR);
		hr = XMLDoc->lpVtbl->Release(XMLDoc);
		CoUninitialize();
		return FALSE;
	}

	// Save XMLDoc to file
	VARIANT varFilename;
	VariantInit(&varFilename);
    varFilename = VariantString(filename);
	hr = XMLDoc->lpVtbl->save(XMLDoc,varFilename);
	hr = XMLDoc->lpVtbl->Release(XMLDoc);
	VariantClear(&varFilename);

	CoUninitialize();
	return SUCCEEDED(hr);
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       get a DOM Document interface                                   //
// Input:         none                                                           //
//                                                                               //
// Output:        pointer to IXMLDOMDocument interface                           //
// Errors:        If the function succeeds, the return value is IXMLDOMDocument  //
//                If the function fails, the return value is NULL.               //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static IXMLDOMDocument * _CreateDocument(void)
{
	HRESULT hr;
	IXMLDOMDocument *pXMLDoc = NULL;

	hr = CoCreateInstance (&CLSID_DOMDocument,       	// CLSID of coclass
							NULL,						// not used - aggregation
							CLSCTX_INPROC_SERVER,		// type of server
							&IID_IXMLDOMDocument,		// IID of interface
							(void**) &pXMLDoc );		// Pointer to our interface pointer

	//////////////////////////////////////////////////////
	// Theese 3 calls shouldn't fail                    //
	//////////////////////////////////////////////////////
	hr = pXMLDoc->lpVtbl->put_async(pXMLDoc,VARIANT_FALSE);
	hr = pXMLDoc->lpVtbl->put_validateOnParse(pXMLDoc,VARIANT_FALSE);
	hr = pXMLDoc->lpVtbl->put_resolveExternals(pXMLDoc,VARIANT_FALSE);
	hr = pXMLDoc->lpVtbl->put_preserveWhiteSpace(pXMLDoc,VARIANT_TRUE);

	if (SUCCEEDED(hr))
	{
		return pXMLDoc;
	}
	// Clean up
    else if (pXMLDoc)
    {
        pXMLDoc->lpVtbl->Release(pXMLDoc);
		CoUninitialize();
    }
    return NULL;
}

///////////////////////////////////////////////////////////////////////////////////
// _AddPI                                                                        //
// Adds a processing instrucction to the IXMLDOMDocument                         //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _AddPI(IXMLDOMDocument *xmlDoc)
{

    IXMLDOMProcessingInstruction *XMLPi = NULL;

    BSTR xml = SysAllocString(L"xml");
    BSTR ver = SysAllocString(L"version=\"1.0\" encoding=\"UTF-8\"");

    HRESULT hr = xmlDoc->lpVtbl->createProcessingInstruction(xmlDoc, xml, ver, &XMLPi );
    if(FAILED(hr))
        goto clean;

    hr = xmlDoc->lpVtbl->appendChild(xmlDoc, (void *)XMLPi, NULL);
    XMLPi->lpVtbl->Release(XMLPi);
    XMLPi = NULL;

clean:
    if(xml)
        SysFreeString(xml);

    if(ver)
        SysFreeString(ver);

    return SUCCEEDED(hr);
}

///////////////////////////////////////////////////////////////////////////////////
// _AddComment                                                                   //
// Adds a comment to the IXMLDOMDocument                                         //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _AddComment(IXMLDOMDocument *xmlDoc, TCHAR *comment)
{

    //CString convert( comment );
    //BSTR bstrComment = convert.AllocSysString();

    IXMLDOMComment *commentNode = NULL;
#ifndef _UNICODE
	BSTR bstrComment = BSTRtoANSI(nodeName);
    HRESULT hr = xmlDoc->lpVtbl->createComment(xmlDoc, comment, &bstrComment);
	SysFreeString(bstrComment);
#else
	HRESULT hr = xmlDoc->lpVtbl->createComment(xmlDoc, comment, &commentNode);
#endif

    if( FAILED(hr) )
        return FALSE;

    hr = xmlDoc->lpVtbl->appendChild(xmlDoc, (void *)commentNode, NULL );
    commentNode->lpVtbl->Release(commentNode);
    commentNode = NULL;

    return SUCCEEDED(hr);
}

///////////////////////////////////////////////////////////////////////////////////
// _AddChild                                                                     //
// Adds a child to the IXMLDOMDocument                                           //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _AddChild(IXMLDOMDocument *xmlDoc, IXMLDOMNode *child)
{
	HRESULT hr = xmlDoc->lpVtbl->appendChild(xmlDoc, child, NULL);
	child->lpVtbl->Release(child);
	return SUCCEEDED(hr);
}

///////////////////////////////////////////////////////////////////////////////////
// _AddRootElement                                                               //
// Adds a root element to the IXMLDOMDocument                                    //
///////////////////////////////////////////////////////////////////////////////////
static IXMLDOMElement * _AddRootElement(IXMLDOMDocument *xmlDoc)
{
    IXMLDOMElement *rootElement = NULL;
	BSTR rootStr = SysAllocString(L"ROOT");
	HRESULT hr = xmlDoc->lpVtbl->createElement(xmlDoc, rootStr, &rootElement);
	SysFreeString(rootStr);

    if( FAILED(hr) )
        return NULL;

    xmlDoc->lpVtbl->appendChild(xmlDoc, (void *)rootElement, NULL );
    return rootElement;
}


///////////////////////////////////////////////////////////////////////////////////
// Purpose:       retrieve IXMLDOMDocument from the XML file                     //
// Input:         TCHAR *fileName:                                               //
//                                                                               //
// Output:        IXMLDOMDocument* pointer.                                      //
// Errors:        If the function succeeds, the return value is IXMLDOMDocument* //
//                If the function fails, the return value is the error code.     //
// Notes:		  This function can only handle single window.                   //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static IXMLDOMDocument* _GetDomDocument(TCHAR *fileName)
{
	HRESULT hr;
	IXMLDOMDocument   *pXMLDoc	= NULL;
	IXMLDOMParseError *pXMLErr	= NULL;
	BSTR bstr 					= NULL;

	VARIANT_BOOL status;
	VARIANT var;

	//////////////////////////////////////////////////////
	// Instance the COM object                          //
	//////////////////////////////////////////////////////
	hr = CoCreateInstance (&CLSID_DOMDocument,       	// CLSID of coclass
							NULL,						// not used - aggregation
							CLSCTX_INPROC_SERVER,		// type of server
							&IID_IXMLDOMDocument,		// IID of interface
							(void**) &pXMLDoc );		// Pointer to our interface pointer
	if (SUCCEEDED(hr))
	{	
		//////////////////////////////////////////////////////
		// Theese 3 calls shouldn't fail                    //
		//////////////////////////////////////////////////////
		pXMLDoc->lpVtbl->put_async(pXMLDoc,VARIANT_FALSE);
		pXMLDoc->lpVtbl->put_validateOnParse(pXMLDoc,VARIANT_FALSE);
		pXMLDoc->lpVtbl->put_resolveExternals(pXMLDoc,VARIANT_FALSE);

#ifndef _UNICODE
		VariantInit(&var);
		V_BSTR(&var) = ANSItoBSTR(fileName);
		V_VT(&var) = VT_BSTR;
#else
		//V_BSTR(&var) = fileName; SysAllocString(
		var = VariantString(fileName);
#endif
		hr = pXMLDoc->lpVtbl->load(pXMLDoc,var, &status);
		VariantClear(&var);
		//////////////////////////////////////////////////////
		// Get error messages                               //
		//////////////////////////////////////////////////////
		if (status != VARIANT_TRUE)
		{
			pXMLDoc->lpVtbl->get_parseError(pXMLDoc,&pXMLErr);
			pXMLErr->lpVtbl->get_reason(pXMLErr,&bstr);
#ifndef _UNICODE
			ansiStr = BSTRtoANSI(bstr);
			MessageBox(0, ansiStr,_T("parseError() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
			free(ansiStr);

#else
			MessageBox(0, bstr,_T("parseError() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
#endif
			SysFreeString(bstr);
			goto cleanup;
		}
	}
	else
	{
		LPTSTR errMsg = _ErrorString(hr);
		MessageBox(0,errMsg,_T("Error in CoCreateInstance() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
		LocalFree(errMsg);
		goto cleanup;
	}

	//////////////////////////////////////////////////////
	// Exit no need to release the interface            //
	//////////////////////////////////////////////////////
	//if (bstr) SysFreeString(bstr);
	//if (&var) VariantClear(&var);
	//if (ansistr) free(ansistr);
	return pXMLDoc;
cleanup:
	//if (bstr) SysFreeString(bstr);
	//if (&var) VariantClear(&var);
	//if (ansistr) free(ansistr);
	if (&var) VariantClear(&var);
	pXMLDoc->lpVtbl->Release(pXMLDoc);
	pXMLErr->lpVtbl->Release(pXMLErr);
	return (void *)NULL;
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       retrieve IXMLDOMElement from IXMLDOMDocument                   //
// Input:         IXMLDOMDocument *pXMLDoc                                       //
//                                                                               //
// Output:        IXMLDOMElement* pointer.                                       //
// Errors:        If the function succeeds, the return value is IXMLDOMElement*  //
//                If the function fails, the return value is NULL.               //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static IXMLDOMElement* _GetRootElement(IXMLDOMDocument *pXMLDoc)
{
	HRESULT hr;
	IXMLDOMElement* pRootElem	= NULL;

	hr = pXMLDoc->lpVtbl->get_documentElement(pXMLDoc,&pRootElem);

	if SUCCEEDED(hr)
	{
		return pRootElem;
	}
	else
	{
		pRootElem->lpVtbl->Release(pRootElem);
		LPTSTR errMsg = _ErrorString(hr);
		MessageBox(0,errMsg,_T("_GetRootElement() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
		LocalFree(errMsg);
	}
	return NULL;
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       retrieve IXMLDOMNodeList* from IXMLDOMElement*                 //
// Input:         IXMLDOMElement* domElement: parent element                     //
//                                                                               //
// Output:        IXMLDOMNodeList* pointer.                                      //
// Errors:        If the function succeeds, the return value is IXMLDOMNodeList* //
//                If the function fails, the return value is NULL.               //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static IXMLDOMNodeList* _GetChildNodes(IXMLDOMElement *domElement)
{
	HRESULT hr;
	IXMLDOMNodeList *childList	= NULL;

	hr = domElement->lpVtbl->get_childNodes(domElement,&childList);
	if SUCCEEDED(hr)
	{
		return childList;
	}
	else
	{
		childList->lpVtbl->Release(childList);
		LPTSTR errMsg = _ErrorString(hr);
		MessageBox(0,errMsg,_T("_GetChildNodes() !"),MB_OK | MB_ICONSTOP | MB_SETFOREGROUND);
		LocalFree(errMsg);
		return NULL;
	}
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Fill tree from starting element deep down 1 level only         //
// Input:         IXMLDOMElement *node:  starting point                          //
// Input:         HTREEITEM hParent: Position in the tree from where to insert   //
// Input:         HTREEITEM hInsAfter: member of TV_INSERTSTRUCT usually 0       //
//                                                                               //
// Output:        True or False on success.                                      //
// Errors:        If the function succeeds, the return value is TRUE             //
// Notes:		  This function can only handle single window.                   //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _PopulateNode(IXMLDOMElement *node,HTREEITEM hParent, HTREEITEM hInsAfter,HWND tree)
{
	HRESULT             hr;
	BSTR				nodeName = NULL;
	HTREEITEM			newParent = NULL;
	DOMNodeType			nodeType;
	IXMLDOMNodeList		*pXMLNodeList;
	long				lCount;

	node->lpVtbl->get_childNodes(node, &pXMLNodeList);
	pXMLNodeList->lpVtbl->get_length(pXMLNodeList, &lCount);
	for(long i = 0; i < lCount; i++) 
	{
		IXMLDOMNode* nodeItem = NULL;
		pXMLNodeList->lpVtbl->get_item(pXMLNodeList,i, &nodeItem);
		nodeItem->lpVtbl->get_nodeType(nodeItem,&nodeType);

		switch(nodeType)
		{
			case NODE_ELEMENT:
			{
				nodeItem->lpVtbl->get_nodeName(nodeItem,&nodeName);
#ifndef _UNICODE
				char *tmpStr;
				tmpStr = BSTRtoANSI(nodeName);
				newParent = InsertTreeItem(hParent,tmpStr,hInsAfter,1,tree,(IXMLDOMElement *)nodeItem);
				free(tmpStr);
#else
				newParent = InsertTreeItem(hParent,nodeName,hInsAfter,1,tree,(IXMLDOMElement *)nodeItem);
#endif
				SysFreeString(nodeName);
				/////////////////////////////////////////////////////
				// I did not want to put attributes at the same level
				// As the childs because attributes are not in this
				// category, but to add visibility to attributes
				// Uncomment next line
				//_PopulateAttributes((IXMLDOMElement *)nodeItem,newParent,hInsAfter,tree);
				break;
			}
			case NODE_TEXT:
			{
				BSTR nodeText;
				nodeItem->lpVtbl->get_text(nodeItem,&nodeText);
#ifndef _UNICODE
				char *tmpStr;
				tmpStr = BSTRtoANSI(nodeText);
				newParent = InsertTreeItem(hParent,tmpStr,hInsAfter,2,tree,(IXMLDOMElement *)nodeItem);
				free(tmpStr);
#else
				newParent = InsertTreeItem(hParent,nodeText,hInsAfter,2,tree,(IXMLDOMElement *)nodeItem);
#endif
				SysFreeString(nodeText);
				break;
			}
			case NODE_CDATA_SECTION:
			{
				hr = nodeItem->lpVtbl->get_nodeName(nodeItem,&nodeName);
#ifndef _UNICODE
				char *tmpStr;
				tmpStr = BSTRtoANSI(nodeName);
				newParent = InsertTreeItem(hParent,tmpStr,hInsAfter,4,tree,(IXMLDOMElement *)nodeItem);
				free(tmpStr);				
#else
				newParent = InsertTreeItem(hParent,nodeName,hInsAfter,4,tree,(IXMLDOMElement *)nodeItem);
#endif
				SysFreeString(nodeName);
				break;
			
			}
			case NODE_COMMENT:
			{
				hr = nodeItem->lpVtbl->get_nodeName(nodeItem,&nodeName);
#ifndef _UNICODE
				char *tmpStr;
				tmpStr = BSTRtoANSI(nodeName);
				newParent = InsertTreeItem(hParent,tmpStr,hInsAfter,3,tree,(IXMLDOMElement *)nodeItem);
				free(tmpStr);
#else
				newParent = InsertTreeItem(hParent,nodeName,hInsAfter,3,tree,(IXMLDOMElement *)nodeItem);
#endif
				SysFreeString(nodeName);
				break;
			}
			default:
			{
				hr = nodeItem->lpVtbl->get_nodeName(nodeItem,&nodeName);
#ifndef _UNICODE
				char *tmpStr;
				tmpStr = BSTRtoANSI(nodeName);
				newParent = InsertTreeItem(hParent,tmpStr,hInsAfter,4,tree,(IXMLDOMElement *)nodeItem);
				free(tmpStr);				
#else
				newParent = InsertTreeItem(hParent,nodeName,hInsAfter,4,tree,(IXMLDOMElement *)nodeItem);
#endif
				SysFreeString(nodeName);
				break;
			}
		}
	}
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Get node attributes                                            //
// Input:         IXMLDOMElement (input node):                                   //
//                HTREEITEM (Parent to insert the item)                          //
//                HTREEITEM (Where after)                                        //
//                HWND (Handle of the tree control)                              //
//                                                                               //
// Output:        BOOL                                                           //
// Errors:        If the function succeeds, the return value is TRUE             //
//                If the function fails, the return value is FALSE               //
// Notes:		  This function can only handle single window.                   //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _PopulateAttributes(IXMLDOMElement *node, HTREEITEM hParent,HTREEITEM hInsAfter,HWND tree)
{
	HRESULT 		hr = S_OK;
	IXMLDOMNamedNodeMap* namedNodeMap = NULL;
	
	hr = node->lpVtbl->get_attributes(node,&namedNodeMap);

	if (SUCCEEDED(hr) && namedNodeMap != NULL) 
	{
		long listLength;
		hr = namedNodeMap->lpVtbl->get_length(namedNodeMap,&listLength);
		for(long i = 0; i < listLength; i++) 
		{
			BSTR nodeName = NULL;
			IXMLDOMNode* listItem = NULL;

			namedNodeMap->lpVtbl->get_item(namedNodeMap,i, &listItem);
			listItem->lpVtbl->get_nodeName(listItem,&nodeName);
#ifndef _UNICODE
			char *	tmpStr;
			tmpStr = BSTRtoANSI(nodeName);
			InsertTreeItem(hParent,tmpStr,hInsAfter,4,tree,(IXMLDOMElement *)listItem);
			free(tmpStr);
#else
			InsertTreeItem(hParent,nodeName,hInsAfter,4,tree,(IXMLDOMElement *)listItem);
#endif
			SysFreeString(nodeName);
		}
	}

clean:
	namedNodeMap->lpVtbl->Release(namedNodeMap);
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////////
// Purpose:       Insert new items in the tree control                           //
// Input:         hParent (parent tree item)                                     //
//                szText (caption of the item)                                   //
//                hInsAfter (insert after item)                                  //
//                iImage (image index from image list)                           //
//                hwndTree (Handle of the SysTreeView32 control)                 //
//                                                                               //
//                                                                               //
// Output:        HWND of the created window                                     //
// Errors:        If the function succeeds, the return value is treeview         //
//                If the function fails, the return value is NULL.               //
//                COM may trow exceptions                                        //
///////////////////////////////////////////////////////////////////////////////////
HTREEITEM InsertTreeItem(HTREEITEM hParent, TCHAR *szText,HTREEITEM hInsAfter,int iImage, HWND hwndTree, IXMLDOMElement* node)
{
	HRESULT             hr;
	HTREEITEM			hItem;
	TV_ITEM				tvI;
	TV_INSERTSTRUCT		tvIns;
	
	tvI.mask			= TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
	tvI.pszText			= szText;
	tvI.cchTextMax		= lstrlen(szText);
	tvI.iImage			= iImage;
	tvI.iSelectedImage	= iImage;
	
	if(node)
	{
		tvI.lParam	= (DWORD)node;
	}

	tvIns.item			= tvI;
	tvIns.hInsertAfter	= hInsAfter;
	tvIns.hParent		= hParent;

	hItem = (HTREEITEM)SendMessage(hwndTree, TVM_INSERTITEM, 0,(LPARAM)(LPTV_INSERTSTRUCT)&tvIns);

	//////////////////////////////////////////////////////
	// Set button state.                                //
	//////////////////////////////////////////////////////
	if (node)
	{
		VARIANT_BOOL hasNodes;
		node->lpVtbl->hasChildNodes(node,&hasNodes);
		if (hasNodes == VARIANT_TRUE)
		{
			TV_ITEM				tvChildItem;
			TV_INSERTSTRUCT		tvChildIns;
				
			tvChildItem.mask			= TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
			tvChildItem.pszText			= _T("");
			tvChildItem.cchTextMax		= 0;
			tvChildItem.iImage			= iImage;
			tvChildItem.iSelectedImage	= iImage;
			tvChildItem.lParam			= (DWORD)NULL;
			tvChildIns.item				= tvChildItem;
			tvChildIns.hInsertAfter		= 0;
			tvChildIns.hParent			= hItem;

			SendMessage(hwndTree, TVM_INSERTITEM, 0,(LPARAM)(LPTV_INSERTSTRUCT)&tvChildIns);
		}
		else
		{
			DOMNodeType nodeType;
			node->lpVtbl->get_nodeType(node,&nodeType);
			if(nodeType == NODE_COMMENT || nodeType == NODE_CDATA_SECTION)
			{
				TV_ITEM				tvChildItem;
				TV_INSERTSTRUCT		tvChildIns;
				
				tvChildItem.mask			= TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
				tvChildItem.pszText			= _T("");
				tvChildItem.cchTextMax		= 0;
				tvChildItem.iImage			= iImage;
				tvChildItem.iSelectedImage	= iImage;
				tvChildItem.lParam	= (DWORD)NULL;

				tvChildIns.item			= tvChildItem;
				tvChildIns.hInsertAfter	= 0;
				tvChildIns.hParent		= hItem;
				//HTREEITEM hChildItem = (HTREEITEM)SendMessage(hwndTree, TVM_INSERTITEM, 0,(LPARAM)(LPTV_INSERTSTRUCT)&tvChildIns);
			}
		}
	}

clean:
	return (hItem);
}

///////////////////////////////////////////////////////////////////////////////////
// Set the imageList to the treeView control                                     //
///////////////////////////////////////////////////////////////////////////////////
static BOOL _SetTreeImages(HWND treeHwd, HINSTANCE hInst)
{
	HICON 		hIcon = NULL;				// The handle icon.
	HIMAGELIST 	hImageList = NULL;			// The handle to the image list.
	if(!hInst)
	hInst = GetModuleHandle(NULL);			// Needs to work in multi thread enviroment ?

	hImageList = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), ILC_COLOR32|ILC_MASK, 0, 0); // Load Image list
	//////////////////////////////////////////////////////
	// Load the resources                               //
	//////////////////////////////////////////////////////
	hIcon = LoadImage(hInst, MAKEINTRESOURCE(IDI_ROOT), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_ELEMENT), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);
	hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_TEXT), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);	
	hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_ATTRIBUTE), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);	
	hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_COMENT), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);	
	hIcon = (HICON)LoadImage(hInst, MAKEINTRESOURCE(IDI_DATA), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	ImageList_AddIcon(hImageList, hIcon);	
	//////////////////////////////////////////////////////
	// Assign the images                                //
	//////////////////////////////////////////////////////
	TreeView_SetImageList(treeHwd, hImageList, TVSIL_NORMAL);
	
	return TRUE;
}

///////////////////////////////////////////////////////////////////////////////////
// BSTRtoANSI()                                                                  //
// Converts a BSTR to ANSI string remember you need to free() the string         //
///////////////////////////////////////////////////////////////////////////////////
static char * BSTRtoANSI(BSTR bstr)
{
	char * ansistr;
	int lenWIDE = SysStringLen(bstr);
	int lenANSI = WideCharToMultiByte(CP_ACP, 0, bstr, lenWIDE, 0, 0, NULL, NULL);
	if (lenANSI > 0)
	{
		ansistr = malloc(lenANSI + 1);
		WideCharToMultiByte(CP_ACP, 0, bstr, lenWIDE, ansistr, lenANSI, NULL, NULL);
		ansistr[lenANSI] = 0;
	}
	else
	{
		//////////////////////////////////////////////////////
		// TO-DO: handle the error                          //
		//////////////////////////////////////////////////////
		return NULL; 
	}
	return ansistr;
}

///////////////////////////////////////////////////////////////////////////////////
// ANSItoBSTR                                                                    //
// Converts a ANSI to BSTR string remember you need to SysFreeString();          //
///////////////////////////////////////////////////////////////////////////////////
static BSTR ANSItoBSTR(char *ansistr)
{
	int lenANSI = lstrlenA(ansistr);
	int lenWIDE;
	BSTR bstr;

	lenWIDE = MultiByteToWideChar(CP_ACP, 0, ansistr, lenANSI, 0, 0);
	if (lenWIDE > 0)
	{
		//////////////////////////////////////////////////////
		// Check whether conversion was successful          //
		//////////////////////////////////////////////////////
		bstr = SysAllocStringLen(0, lenWIDE);
		MultiByteToWideChar(CP_ACP, 0, ansistr, lenANSI, bstr, lenWIDE);
	}
	else
	{
		//////////////////////////////////////////////////////
		// TO-DO handle the error                           //
		//////////////////////////////////////////////////////
		return NULL;
	}		
	return bstr;
}

///////////////////////////////////////////////////////////////////////////////////
// VariantString                                                                 //
// Converts a BSTR to Variant string remember you need to VariantClear();        //
///////////////////////////////////////////////////////////////////////////////////
static VARIANT VariantString(BSTR str)
{
    VARIANT var;
    VariantInit(&var);
    V_BSTR(&var) = SysAllocString(str);
    V_VT(&var) = VT_BSTR;
    return var;
}

///////////////////////////////////////////////////////////////////////////////////
// Get extended information from error codes                                     //
// FORMAT_MESSAGE_ALLOCATE_BUFFER FormatMessage will allocate the buffer         //
///////////////////////////////////////////////////////////////////////////////////
static TCHAR * _ErrorString(DWORD errno)
{

	TCHAR *szMsgBuf;

	FormatMessage(
			FORMAT_MESSAGE_FROM_SYSTEM|FORMAT_MESSAGE_ALLOCATE_BUFFER
		    , NULL
		    , errno
		    , MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT)
		    , (LPTSTR) &szMsgBuf
		    , 0x1000 * sizeof(TCHAR)
		    , NULL);

	return szMsgBuf;
}

// End of XMLSerialTree.c

