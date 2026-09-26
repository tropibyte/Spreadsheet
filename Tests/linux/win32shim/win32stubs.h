#pragma once
// Test hook into the Win32 stubs: every SendMessage the code under test makes
// is recorded here, so a test can assert that a notification was sent without
// standing up a real window to receive it.
#include <windows.h>
#include <vector>

struct StubSentMessage {
    HWND   hwnd;
    UINT   msg;
    WPARAM wParam;
    LPARAM lParam;
    UINT   notifyCode;   // NMHDR::code, captured during the call for WM_NOTIFY
};

std::vector<StubSentMessage>& StubMessageLog();
void StubMessageLogClear();
// Number of WM_NOTIFY messages carrying `code` since the last clear.
size_t StubCountNotifications(UINT code);
