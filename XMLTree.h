///////////////////////////////////////////////////////////////////////////////////
//                                                                               //
//                       THIS IS THE RELEASE VERSION                             //
//                                                                               //
///////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////
// File Name:           XMLTree.h                                                //
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
///////////////////////////////////////////////////////////////////////////////////

#ifndef __XMLTREE_H__
#define __XMLTREE_H__

///////////////////////////////////////////////////////////////////////////////////
//                                                                               //
// External visible functions                                                    //
//                                                                               //
///////////////////////////////////////////////////////////////////////////////////
BOOL		PopulateTree	(HWND tree, TCHAR *fileName);
BOOL        ClearTree       (HWND hwndTree);
HTREEITEM	InsertTreeItem	(HTREEITEM hParent, TCHAR *szText,HTREEITEM hInsAfter,int iImage, HWND hwndTree, IXMLDOMElement* node);
BOOL		SaveTree		(HWND hWnd, TCHAR *filename);

// Message Reflect
int			OnNotify        (HWND hWnd, UINT message, int idCtrl, NMHDR *pnmh);
int         OnMouseMove     (HWND hWnd, WPARAM wParam, int xPos, int yPos);
int			OnLButtonDown   (HWND hWnd, WPARAM wParam, int xPos, int yPos);
int			OnLButtonUp     (HWND hWnd, WPARAM wParam, int xPos, int yPos);
void        OnBegindragTree (NMHDR* pnmh, LRESULT* pResult);


#endif

// End of XMLSerialTree.h
