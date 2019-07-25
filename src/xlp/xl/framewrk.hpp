///***************************************************************************
// File:        FRAMEWRK.H
//
// Purpose:        Header file for Framework library
//
// Platform:    Microsoft Windows
//
// Comments:
//              Include this file in any source files
//              that use the framework library.
//
// From the Microsoft Excel Developer's Kit, Version 8
// Copyright (c) 1997 Microsoft Corporation. All rights reserved.
///***************************************************************************
#ifndef xlsdk_framewrk_hpp
#define xlsdk_framewrk_hpp

#include <string>
//
// Total amount of memory to allocate for all temporary XLOPERs
//

#define MEMORYSIZE 1024


// 
// Function prototypes
//

void far __cdecl debugPrintf(LPSTR lpFormat, ...);
LPSTR GetTempMemory(int cBytes);
void FreeAllTempMemory(void);
void __cdecl Excel(int xlfn, LPXLOPER pxResult, int count, ...);
LPXLOPER TempNum(double d);
//LPXLOPER TempStr(LPSTR lpstr);
LPXLOPER TempStrNoSize(LPSTR lpstr);
LPXLOPER TempStrStl(const std::string &s);
LPXLOPER TempBool(int b);
LPXLOPER TempInt(short int i);
LPXLOPER TempActiveRef(WORD rwFirst,WORD rwLast,BYTE colFirst,BYTE colLast);
LPXLOPER TempActiveCell(WORD rw, BYTE col);
LPXLOPER TempActiveRow(WORD rw);
LPXLOPER TempActiveColumn(BYTE col);
LPXLOPER TempErr(WORD i);
LPXLOPER TempMissing(void);
LPXLOPER TempNil();

#endif
