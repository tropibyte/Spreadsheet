// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"

HINSTANCE g_hInstDLL;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    {
        DisableThreadLibraryCalls(hModule);
        InitializeGridLibrary(hModule);
        g_hInstDLL = hModule;
    }
    break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
        break;
    case DLL_PROCESS_DETACH:
        UnInitializeGridLibrary(hModule);
        break;
    }
    return TRUE;
}

