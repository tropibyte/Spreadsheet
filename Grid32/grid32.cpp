#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"


bool RegisterGrid32Class(HINSTANCE hInstance)
{
    WNDCLASSEX wc = { 0 };

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS | CS_OWNDC | CS_GLOBALCLASS;
    wc.lpfnWndProc = CGrid32Mgr::Grid32_WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 2 * sizeof(PVOID);
    wc.hInstance = hInstance;
    wc.hIcon = NULL;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = GRID_WNDCLASS_NAME;
    wc.hIconSm = NULL;

    return RegisterClassEx(&wc) != NULL;
}

bool InitializeGridLibrary(HMODULE hModule)
{
    if (!RegisterGrid32Class(hModule))
        return false;

    return true;
}

bool UnInitializeGridLibrary(HMODULE hModule)
{
    return UnregisterClass(GRID_WNDCLASS_NAME, hModule) != NULL;
}
