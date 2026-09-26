#pragma once
// Declarations only — this shim exists so the real Grid32 sources can be
// type-checked on a POSIX host. Nothing here is ever executed.

// ---- constants -------------------------------------------------------------
#define RGB(r,g,b) ((COLORREF)(((BYTE)(r))|((WORD)((BYTE)(g))<<8)|(((DWORD)(BYTE)(b))<<16)))
#define GetRValue(c) ((BYTE)(c))
#define GetGValue(c) ((BYTE)(((WORD)(c))>>8))
#define GetBValue(c) ((BYTE)((c)>>16))
#define LOWORD(l) ((WORD)(((DWORD_PTR)(l)) & 0xffff))
#define HIWORD(l) ((WORD)((((DWORD_PTR)(l))>>16) & 0xffff))
#define MAKELPARAM(l,h) ((LPARAM)(((WORD)(l))|(((DWORD)((WORD)(h)))<<16)))
#define MAKEWPARAM(l,h) ((WPARAM)(((WORD)(l))|(((DWORD)((WORD)(h)))<<16)))
#define MAKELRESULT(l,h) ((LRESULT)(((WORD)(l))|(((DWORD)((WORD)(h)))<<16)))

enum {
  WM_NULL=0, WM_CREATE=1, WM_DESTROY=2, WM_MOVE=3, WM_SIZE=5, WM_SETFOCUS=7,
  WM_KILLFOCUS=8, WM_PAINT=15, WM_ERASEBKGND=20, WM_SHOWWINDOW=24,
  WM_SETFONT=48, WM_GETFONT=49, WM_NCCREATE=129, WM_NCDESTROY=130,
  WM_NCHITTEST=132, WM_NCPAINT=133, WM_NCMOUSEMOVE=160, WM_KEYDOWN=256,
  WM_KEYUP=257, WM_CHAR=258, WM_SYSKEYDOWN=260, WM_SYSKEYUP=261,
  WM_TIMER=275, WM_HSCROLL=276, WM_VSCROLL=277, WM_MOUSEMOVE=512,
  WM_LBUTTONDOWN=513, WM_LBUTTONUP=514, WM_RBUTTONDOWN=516, WM_RBUTTONUP=517,
  WM_MOUSEWHEEL=522, WM_MOUSELEAVE=675, WM_MOUSEHOVER=673,
  WM_CAPTURECHANGED=533, WM_CANCELMODE=31, WM_SETREDRAW=11,
  WM_CUT=768, WM_COPY=769, WM_PASTE=770, WM_CLEAR=771, WM_UNDO=772,
  WM_USER=1024, WM_NOTIFY=78, WM_COMMAND=273,
  EM_SETSEL=177, EM_SCROLLCARET=183, EM_GETSEL=176,
  ES_AUTOHSCROLL=128, ES_MULTILINE=4,
  WS_CHILD=0x40000000, WS_VISIBLE=0x10000000, WS_BORDER=0x00800000,
  WS_VSCROLL=0x00200000, WS_HSCROLL=0x00100000,
  SW_HIDE=0, SW_SHOW=5,
  SB_HORZ=0, SB_VERT=1, SB_LINEUP=0, SB_LINEDOWN=1, SB_PAGEUP=2,
  SB_PAGEDOWN=3, SB_THUMBPOSITION=4, SB_THUMBTRACK=5, SB_TOP=6, SB_BOTTOM=7,
  SB_ENDSCROLL=8, SB_LINELEFT=0, SB_LINERIGHT=1, SB_PAGELEFT=2,
  SB_PAGERIGHT=3, SB_LEFT=6, SB_RIGHT=7,
  SIF_RANGE=1, SIF_PAGE=2, SIF_POS=4, SIF_TRACKPOS=16, SIF_ALL=23,
  DT_LEFT=0, DT_CENTER=1, DT_RIGHT=2, DT_TOP=0, DT_VCENTER=4, DT_BOTTOM=8,
  DT_WORDBREAK=16, DT_SINGLELINE=32, DT_NOCLIP=256, DT_CALCRECT=1024,
  PS_SOLID=0, PS_DASH=1, PS_DOT=2, PS_NULL=5,
  TRANSPARENT=1, OPAQUE=2, SRCCOPY=0x00CC0020,
  R2_NOT=6, R2_COPYPEN=13, R2_XORPEN=7,
  CF_UNICODETEXT=13, GMEM_MOVEABLE=2, GMEM_ZEROINIT=64,
  GWLP_WNDPROC=-4, GWLP_ID=-12, GWL_STYLE=-16,
  HTNOWHERE=0, HTCLIENT=1, HTCAPTION=2, HTHELP=21,
  NM_FIRST=0, NM_HOVER=-13, NM_RCLICK=-5, NM_RDBLCLK=-6, NM_SETFOCUS=-7,
  SPI_GETMOUSEHOVERTIME=102, NM_OUTOFMEMORY=-1, NM_CLICK=-2, NM_DBLCLK=-3, NM_KILLFOCUS=-8,
  TME_LEAVE=2, TME_HOVER=1, TME_CANCEL=0x80000000,
  MB_ICONERROR=16, MB_OK=0,
  LOGPIXELSX=88, LOGPIXELSY=90, HORZRES=8, VERTRES=10,
  WHITE_BRUSH=0, NULL_BRUSH=5, PS_DASHDOT=3, _TRUNCATE=(size_t)-1, BLACK_PEN=7, DEFAULT_GUI_FONT=17, SYSTEM_FONT=13,
  VK_BACK=8, VK_TAB=9, VK_RETURN=13, VK_SHIFT=16, VK_CONTROL=17,
  VK_ESCAPE=27, VK_SPACE=32, VK_PRIOR=33, VK_NEXT=34, VK_END=35,
  VK_HOME=36, VK_LEFT=37, VK_UP=38, VK_RIGHT=39, VK_DOWN=40,
  VK_DELETE=46, VK_F1=112,
  WHEEL_DELTA=120, HOVER_DEFAULT=0xFFFFFFFF,
  IDC_ARROW_=32512
};
#define IDC_ARROW    ((LPCWSTR)(intptr_t)32512)
#define IDC_SIZENS   ((LPCWSTR)(intptr_t)32645)
#define IDC_SIZEWE   ((LPCWSTR)(intptr_t)32644)
#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
#define GET_WHEEL_DELTA_WPARAM(w) ((short)HIWORD(w))
#define _ASSERT(x) ((void)0)

// ---- structs ---------------------------------------------------------------
struct SCROLLINFO { UINT cbSize, fMask; int nMin, nMax; UINT nPage; int nPos, nTrackPos; };
struct TRACKMOUSEEVENT { DWORD cbSize, dwFlags; HWND hwndTrack; DWORD dwHoverTime; };
struct CREATESTRUCT { void* lpCreateParams; void* hInstance; HMENU hMenu; HWND hwndParent;
                      int cy, cx, y, x; LONG style; LPCWSTR lpszName, lpszClass; DWORD dwExStyle; };
struct SIZE { LONG cx, cy; };
struct LOGFONTW { LONG lfHeight, lfWidth, lfEscapement, lfOrientation, lfWeight;
                  BYTE lfItalic, lfUnderline, lfStrikeOut, lfCharSet, lfOutPrecision,
                       lfClipPrecision, lfQuality, lfPitchAndFamily; wchar_t lfFaceName[32]; };
typedef LOGFONTW LOGFONT;
typedef void* HCURSOR;
typedef void* HINSTANCE;
typedef void* HGDIOBJ;
typedef SCROLLINFO* LPSCROLLINFO;
typedef TRACKMOUSEEVENT* LPTRACKMOUSEEVENT;

// ---- GDI -------------------------------------------------------------------
HBRUSH CreateSolidBrush(COLORREF);
HPEN   CreatePen(int, int, COLORREF);
HFONT  CreateFontW(int,int,int,int,int,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,LPCWSTR);
HFONT  CreateFontIndirect(const LOGFONT*);
HGDIOBJ SelectObject(HDC, HGDIOBJ);
BOOL   DeleteObject(HGDIOBJ);
HGDIOBJ GetStockObject(int);
HDC    CreateCompatibleDC(HDC);
HBITMAP CreateCompatibleBitmap(HDC, int, int);
BOOL   DeleteDC(HDC);
BOOL   BitBlt(HDC,int,int,int,int,HDC,int,int,DWORD);
int    FillRect(HDC, const RECT*, HBRUSH);
BOOL   Rectangle(HDC,int,int,int,int);
int    DrawTextW(HDC, LPCWSTR, int, RECT*, UINT);
int    SetBkMode(HDC, int);
COLORREF SetTextColor(HDC, COLORREF);
COLORREF SetBkColor(HDC, COLORREF);
int    GetDeviceCaps(HDC, int);
BOOL   MoveToEx(HDC,int,int,POINT*);
BOOL   LineTo(HDC,int,int);
int    SetROP2(HDC,int);
BOOL   GetTextExtentPoint32W(HDC, LPCWSTR, int, SIZE*);
BOOL   PatBlt(HDC,int,int,int,int,DWORD);

// ---- User32 ----------------------------------------------------------------
HDC    GetDC(HWND);
int    ReleaseDC(HWND, HDC);
HDC    BeginPaint(HWND, PAINTSTRUCT*);
BOOL   EndPaint(HWND, const PAINTSTRUCT*);
BOOL   GetClientRect(HWND, RECT*);
BOOL   GetWindowRect(HWND, RECT*);
BOOL   InvalidateRect(HWND, const RECT*, BOOL);
BOOL   UpdateWindow(HWND);
HWND   CreateWindowEx(DWORD, LPCWSTR, LPCWSTR, DWORD, int,int,int,int, HWND, HMENU, HINSTANCE, void*);
BOOL   DestroyWindow(HWND);
BOOL   ShowWindow(HWND, int);
BOOL   MoveWindow(HWND,int,int,int,int,BOOL);
HWND   SetFocus(HWND);
HWND   GetParent(HWND);
BOOL   IsWindow(HWND);
BOOL   BringWindowToTop(HWND);
HCURSOR SetCursor(HCURSOR);
HCURSOR LoadCursor(HINSTANCE, LPCWSTR);
LONG_PTR GetWindowLongPtr(HWND, int);
LONG_PTR SetWindowLongPtr(HWND, int, LONG_PTR);
LRESULT SendMessage(HWND, UINT, WPARAM, LPARAM);
BOOL   PostMessage(HWND, UINT, WPARAM, LPARAM);
LRESULT DefWindowProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CallWindowProc(WNDPROC, HWND, UINT, WPARAM, LPARAM);
BOOL   SetWindowText(HWND, LPCWSTR);
int    GetWindowText(HWND, LPWSTR, int);
int    GetWindowTextLength(HWND);
BOOL   GetCursorPos(POINT*);
BOOL   ScreenToClient(HWND, POINT*);
BOOL   ClientToScreen(HWND, POINT*);
BOOL   ClipCursor(const RECT*);
HWND   SetCapture(HWND);
BOOL   ReleaseCapture();
HWND   GetCapture();
BOOL   TrackMouseEvent(LPTRACKMOUSEEVENT);
UINT_PTR SetTimer(HWND, UINT_PTR, UINT, void*);
BOOL   KillTimer(HWND, UINT_PTR);
short  GetAsyncKeyState(int);
UINT   GetDoubleClickTime();
int    MessageBox(HWND, LPCWSTR, LPCWSTR, UINT);
HWND   GetDesktopWindow();
int    SetScrollInfo(HWND, int, const SCROLLINFO*, BOOL);
BOOL   GetScrollInfo(HWND, int, SCROLLINFO*);
BOOL   InflateRect(RECT*, int, int);
BOOL   OffsetRect(RECT*, int, int);
BOOL   IntersectRect(RECT*, const RECT*, const RECT*);
BOOL   PtInRect(const RECT*, POINT);
BOOL   SetRect(RECT*, int,int,int,int);
BOOL   AlphaBlend(HDC,int,int,int,int,HDC,int,int,int,int,DWORD);

// ---- clipboard / kernel ----------------------------------------------------
BOOL   OpenClipboard(HWND);
BOOL   CloseClipboard();
BOOL   EmptyClipboard();
HANDLE GetClipboardData(UINT);
HANDLE SetClipboardData(UINT, HANDLE);
HGLOBAL GlobalAlloc(UINT, size_t);
void*  GlobalLock(HGLOBAL);
BOOL   GlobalUnlock(HGLOBAL);
HGLOBAL GlobalFree(HGLOBAL);
HMODULE GetModuleHandle(LPCWSTR);
unsigned long long GetTickCount64();

// ---- GDI+ startup ----------------------------------------------------------
namespace Gdiplus {
  struct GdiplusStartupInput { int GdiplusVersion = 1; void* DebugEventCallback = nullptr;
                               BOOL SuppressBackgroundThread = 0; BOOL SuppressExternalCodecs = 0; };
  struct GdiplusStartupOutput { void* NotificationHook; void* NotificationUnhook; };
  int  GdiplusStartup(ULONG_PTR*, const GdiplusStartupInput*, GdiplusStartupOutput*);
  void GdiplusShutdown(ULONG_PTR);
}

int  SetScrollPos(HWND, int, int, BOOL);
int  GetScrollPos(HWND, int);
BOOL SetScrollRange(HWND, int, int, int, BOOL);

typedef POINT* LPPOINT;
BOOL SystemParametersInfo(UINT, UINT, void*, UINT);
int  MapWindowPoints(HWND, HWND, LPPOINT, UINT);

// Windows.h defines min/max as macros; Grid32Mgr.cpp relies on them.
// Functions here so <limits> and friends still compile.
template <class T> inline const T& max(const T& a, const T& b) { return a > b ? a : b; }
template <class T> inline const T& min(const T& a, const T& b) { return a < b ? a : b; }

// Secure-CRT string helpers, as overloads so both the (dst, src) array form
// and the (dst, count, src) form resolve.
template <size_t N> inline int wcscpy_s(wchar_t (&d)[N], const wchar_t* s)
{ wcsncpy(d, s, N - 1); d[N - 1] = 0; return 0; }
inline int wcscpy_s(wchar_t* d, size_t n, const wchar_t* s)
{ if (!n) return 0; wcsncpy(d, s, n - 1); d[n - 1] = 0; return 0; }
inline int _tcsncpy_s(wchar_t* d, size_t n, const wchar_t* s, size_t)
{ if (!n) return 0; wcsncpy(d, s, n - 1); d[n - 1] = 0; return 0; }
