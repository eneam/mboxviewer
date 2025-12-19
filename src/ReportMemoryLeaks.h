//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#pragma once

#define _CRT_SECURE_NO_WARNINGS

#if _DEBUG
/*
How to Apply : #
Option 1 : Visual Studio IDE#
Right - click your project ? Properties.
Navigate to Linker ? Command Line.
In "Additional Options", add /FORCE:MULTIPLE.
Click OK and rebuild.
*/

// Uncomment both and add /FORCE:MULTIPLE as above
//#define _CRTDBG_MAP_ALLOC
//#include <crtdbg.h>

// This creates duplicate functions, follow above to resolve
// Usefull but problematic. Reports this file as source of leaks
// Need to investigate better solution
#ifdef _CRTDBG_MAP_ALLOC
void* operator new(size_t s)
{
    static int cnt = 0;
    cnt++;
    if (cnt == 234094)
        int deb = 1;
    if (s == 232)
        int deb = 1;
    return ::operator new(s, _NORMAL_BLOCK, __FILE__, __LINE__);
}
#endif /* _CRTDBG_MAP_ALLOC */

#if 0
#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif
#endif

#endif

