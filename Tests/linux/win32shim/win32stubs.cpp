// Auto-generated no-op definitions for the Win32 shim declarations.
// Only enough behaviour to let the real Grid32 code run headless.
#include <windows.h>
#include <gdiplus.h>

HBRUSH CreateSolidBrush(COLORREF) { return nullptr; }
HPEN CreatePen(int, int, COLORREF) { return nullptr; }
HFONT CreateFontW(int,int,int,int,int,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,DWORD,LPCWSTR) { return nullptr; }
HFONT CreateFontIndirect(const LOGFONT*) { return nullptr; }
HGDIOBJ SelectObject(HDC, HGDIOBJ) { return nullptr; }
BOOL DeleteObject(HGDIOBJ) { return 0; }
HGDIOBJ GetStockObject(int) { return nullptr; }
HDC CreateCompatibleDC(HDC) { return nullptr; }
HBITMAP CreateCompatibleBitmap(HDC, int, int) { return nullptr; }
BOOL DeleteDC(HDC) { return 0; }
BOOL BitBlt(HDC,int,int,int,int,HDC,int,int,DWORD) { return 0; }
int FillRect(HDC, const RECT*, HBRUSH) { return 0; }
BOOL Rectangle(HDC,int,int,int,int) { return 0; }
int DrawTextW(HDC, LPCWSTR, int, RECT*, UINT) { return 0; }
int SetBkMode(HDC, int) { return 0; }
COLORREF SetTextColor(HDC, COLORREF) { return 0; }
COLORREF SetBkColor(HDC, COLORREF) { return 0; }
int GetDeviceCaps(HDC, int) { return 0; }
BOOL MoveToEx(HDC,int,int,POINT*) { return 0; }
BOOL LineTo(HDC,int,int) { return 0; }
int SetROP2(HDC,int) { return 0; }
BOOL GetTextExtentPoint32W(HDC, LPCWSTR, int, SIZE*) { return 0; }
BOOL PatBlt(HDC,int,int,int,int,DWORD) { return 0; }
HDC GetDC(HWND) { return nullptr; }
int ReleaseDC(HWND, HDC) { return 0; }
HDC BeginPaint(HWND, PAINTSTRUCT*) { return nullptr; }
BOOL EndPaint(HWND, const PAINTSTRUCT*) { return 0; }
BOOL GetClientRect(HWND, RECT*) { return 0; }
BOOL GetWindowRect(HWND, RECT*) { return 0; }
BOOL InvalidateRect(HWND, const RECT*, BOOL) { return 0; }
BOOL UpdateWindow(HWND) { return 0; }
HWND CreateWindowEx(DWORD, LPCWSTR, LPCWSTR, DWORD, int,int,int,int, HWND, HMENU, HINSTANCE, void*) { return nullptr; }
BOOL DestroyWindow(HWND) { return 0; }
BOOL ShowWindow(HWND, int) { return 0; }
BOOL MoveWindow(HWND,int,int,int,int,BOOL) { return 0; }
HWND SetFocus(HWND) { return nullptr; }
HWND GetParent(HWND) { return nullptr; }
BOOL IsWindow(HWND) { return 0; }
BOOL BringWindowToTop(HWND) { return 0; }
HCURSOR SetCursor(HCURSOR) { return nullptr; }
HCURSOR LoadCursor(HINSTANCE, LPCWSTR) { return nullptr; }
LONG_PTR GetWindowLongPtr(HWND, int) { return 0; }
LONG_PTR SetWindowLongPtr(HWND, int, LONG_PTR) { return 1; }
LRESULT SendMessage(HWND, UINT, WPARAM, LPARAM) { return 0; }
BOOL PostMessage(HWND, UINT, WPARAM, LPARAM) { return 0; }
LRESULT DefWindowProc(HWND, UINT, WPARAM, LPARAM) { return 0; }
LRESULT CallWindowProc(WNDPROC, HWND, UINT, WPARAM, LPARAM) { return 0; }
BOOL SetWindowText(HWND, LPCWSTR) { return 0; }
int GetWindowText(HWND, LPWSTR, int) { return 0; }
int GetWindowTextLength(HWND) { return 0; }
BOOL GetCursorPos(POINT*) { return 0; }
BOOL ScreenToClient(HWND, POINT*) { return 0; }
BOOL ClientToScreen(HWND, POINT*) { return 0; }
BOOL ClipCursor(const RECT*) { return 0; }
HWND SetCapture(HWND) { return nullptr; }
BOOL ReleaseCapture() { return 0; }
HWND GetCapture() { return nullptr; }
BOOL TrackMouseEvent(LPTRACKMOUSEEVENT) { return 0; }
UINT_PTR SetTimer(HWND, UINT_PTR, UINT, void*) { return 0; }
BOOL KillTimer(HWND, UINT_PTR) { return 0; }
short GetAsyncKeyState(int) { return 0; }
UINT GetDoubleClickTime() { return 0; }
int MessageBox(HWND, LPCWSTR, LPCWSTR, UINT) { return 0; }
HWND GetDesktopWindow() { return nullptr; }
int SetScrollInfo(HWND, int, const SCROLLINFO*, BOOL) { return 0; }
BOOL GetScrollInfo(HWND, int, SCROLLINFO*) { return 0; }
BOOL InflateRect(RECT*, int, int) { return 0; }
BOOL OffsetRect(RECT*, int, int) { return 0; }
BOOL IntersectRect(RECT*, const RECT*, const RECT*) { return 0; }
BOOL PtInRect(const RECT*, POINT) { return 0; }
BOOL SetRect(RECT*, int,int,int,int) { return 0; }
BOOL AlphaBlend(HDC,int,int,int,int,HDC,int,int,int,int,DWORD) { return 0; }
BOOL OpenClipboard(HWND) { return 0; }
BOOL CloseClipboard() { return 0; }
BOOL EmptyClipboard() { return 0; }
HANDLE GetClipboardData(UINT) { return nullptr; }
HANDLE SetClipboardData(UINT, HANDLE) { return nullptr; }
HGLOBAL GlobalAlloc(UINT, size_t) { return nullptr; }
void* GlobalLock(HGLOBAL) { return nullptr; }
BOOL GlobalUnlock(HGLOBAL) { return 0; }
HGLOBAL GlobalFree(HGLOBAL) { return nullptr; }
HMODULE GetModuleHandle(LPCWSTR) { return nullptr; }
unsigned long long GetTickCount64() { return 0; }
int SetScrollPos(HWND, int, int, BOOL) { return 0; }
int GetScrollPos(HWND, int) { return 0; }
BOOL SetScrollRange(HWND, int, int, int, BOOL) { return 0; }
BOOL SystemParametersInfo(UINT, UINT, void*, UINT) { return 0; }
int MapWindowPoints(HWND, HWND, LPPOINT, UINT) { return 0; }

namespace Gdiplus {
  int  GdiplusStartup(ULONG_PTR* t, const GdiplusStartupInput*, GdiplusStartupOutput*) { if (t) *t = 1; return 0; }
  void GdiplusShutdown(ULONG_PTR) {}
}
