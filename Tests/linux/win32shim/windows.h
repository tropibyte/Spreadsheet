#pragma once
// Minimal Win32 surface so the real Grid32 headers compile on a POSIX host.
// Types only — nothing here is called at runtime by the Formulator tests.
#include <cstddef>
#include <cstring>
#include <cstdio>
#include <cwchar>
#include <cwctype>
#include <cstdint>
typedef int                 BOOL;
typedef int                 INT;
typedef unsigned int        UINT;
typedef unsigned long       DWORD;
typedef long                LONG;
typedef unsigned short      WORD;
typedef unsigned char       BYTE;
typedef unsigned long       COLORREF;
typedef intptr_t            LONG_PTR;
typedef uintptr_t           ULONG_PTR;
typedef uintptr_t           UINT_PTR;
typedef uintptr_t           DWORD_PTR;
typedef uintptr_t           WPARAM;
typedef intptr_t            LPARAM;
typedef intptr_t            LRESULT;
typedef wchar_t             WCHAR, TCHAR;
typedef const wchar_t*      LPCWSTR;
typedef wchar_t*            LPWSTR;
typedef const wchar_t*      LPCTSTR;
typedef void*               HANDLE;
typedef void*               HWND;
typedef void*               HDC;
typedef void*               HFONT;
typedef void*               HPEN;
typedef void*               HBRUSH;
typedef void*               HBITMAP;
typedef void*               HMODULE;
typedef void*               HMENU;
typedef void*               HGLOBAL;
#define CALLBACK
#define WINAPI
#define _T(x) L##x
#define TRUE  1
#define FALSE 0
#ifndef NULL
#define NULL 0
#endif
struct POINT { LONG x, y; };
struct RECT  { LONG left, top, right, bottom; };
struct PAINTSTRUCT { HDC hdc; BOOL fErase; RECT rcPaint; };
struct tagNMHDR { HWND hwndFrom; UINT_PTR idFrom; UINT code; };
typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
#define UNREFERENCED_PARAMETER(P) ((void)(P))
#define swprintf_s swprintf
static inline int _wcsicmp(const wchar_t* a, const wchar_t* b) {
    for (;; ++a, ++b) {
        wchar_t ca = (wchar_t)towlower(*a), cb = (wchar_t)towlower(*b);
        if (ca != cb) return ca < cb ? -1 : 1;
        if (!ca) return 0;
    }
}

#include "win32api.h"
