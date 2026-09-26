#pragma once
#include <windows.h>
namespace Gdiplus {
  enum Status { Ok = 0, GenericError = 1 };
  struct Color {
      Color(BYTE a, BYTE r, BYTE g, BYTE b) { (void)a;(void)r;(void)g;(void)b; }
      Color(BYTE r, BYTE g, BYTE b) { (void)r;(void)g;(void)b; }
  };
  struct Rect { int X=0, Y=0, Width=0, Height=0;
                Rect() {}
                Rect(int x,int y,int w,int h): X(x),Y(y),Width(w),Height(h) {} };
  struct SolidBrush { SolidBrush(const Color&) {} };
  struct Graphics {
      Graphics(HDC) {}
      Status FillRectangle(const SolidBrush*, const Rect&) { return Ok; }
  };
}
using namespace Gdiplus;
