#include <iostream>
#include <windows.h>
#include <uiautomation.h>
#include <string>

// 自动链接必要的库
#pragma comment(lib, "uiautomationcore.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// --- 全局定义 ---
#define HOTKEY_ID 1

// 函数声明
void ExtractTextAndCopyToClipboard();
void SetClipboardText(const std::wstring& text);
BOOL WINAPI CtrlHandler(DWORD dwCtrlType);

// --- 主函数 ---
int main()
{
    setlocale(LC_ALL, "");

    if (!SetConsoleCtrlHandler(CtrlHandler, TRUE)) {
        std::cerr << "[Error] Could not set control handler." << std::endl;
        return 1;
    }

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
        std::cerr << "[Error] Failed to initialize COM." << std::endl;
        return 1;
    }

    if (!RegisterHotKey(NULL, HOTKEY_ID, MOD_CONTROL | MOD_ALT, 0x58))
    {
        std::cerr << "[Error] Failed to register hotkey. It might be in use by another application." << std::endl;
        CoUninitialize();
        return 1;
    }

    std::wcout << L"程序已在后台运行..." << std::endl;
    std::wcout << L"请按 [Ctrl + Alt + X] 来提取鼠标下方的文本。" << std::endl;
    std::wcout << L"请在此窗口中按 [Ctrl + C] 或直接关闭窗口来退出程序。" << std::endl;

    MSG msg = { 0 };
    while (GetMessage(&msg, NULL, 0, 0) != 0)
    {
        if (msg.message == WM_HOTKEY)
        {
            if (msg.wParam == HOTKEY_ID)
            {
                std::wcout << L"\n[Info] Hotkey pressed. Extracting text..." << std::endl;
                ExtractTextAndCopyToClipboard();
            }
        }
    }

    UnregisterHotKey(NULL, HOTKEY_ID);
    CoUninitialize();
    std::wcout << L"\n[Info] Hotkey unregistered. Program exited." << std::endl;

    return 0;
}

void ExtractTextAndCopyToClipboard()
{
    IUIAutomation* pAutomation = NULL;
    IUIAutomationElement* pElement = NULL;
    VARIANT varValue, varName;
    VariantInit(&varValue);
    VariantInit(&varName);
    std::wstring final_text;

    HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&pAutomation);
    if (FAILED(hr) || pAutomation == NULL)
    {
        std::cerr << "[Error] Failed to create UI Automation instance." << std::endl;
        goto cleanup;
    }

    POINT pt;
    if (!GetCursorPos(&pt))
    {
        std::cerr << "[Error] Failed to get cursor position." << std::endl;
        goto cleanup;
    }

    hr = pAutomation->ElementFromPoint(pt, &pElement);
    if (FAILED(hr) || pElement == NULL)
    {
        std::wcout << L"[Info] No UI element found under the cursor." << std::endl;
        goto cleanup;
    }

    pElement->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &varValue);
    pElement->GetCurrentPropertyValue(UIA_NamePropertyId, &varName);

    // 核心逻辑：优先使用Value，若为空，则自动使用Name
    if (varValue.vt == VT_BSTR && varValue.bstrVal != NULL && wcslen(varValue.bstrVal) > 0)
    {
        final_text = varValue.bstrVal;
    }
    else if (varName.vt == VT_BSTR && varName.bstrVal != NULL)
    {
        final_text = varName.bstrVal;
    }

    if (!final_text.empty())
    {
        SetClipboardText(final_text);
        std::wcout << L"[OK] Text extracted and copied to clipboard:" << std::endl;
        std::wcout << final_text << std::endl;
    }
    else
    {
        std::wcout << L"[Info] Found a UI element, but it has no text in its Name or Value properties." << std::endl;
    }

cleanup:
    if (pAutomation) pAutomation->Release();
    if (pElement) pElement->Release();
    VariantClear(&varValue);
    VariantClear(&varName);
}

void SetClipboardText(const std::wstring& text)
{
    if (!OpenClipboard(NULL)) return;

    EmptyClipboard();
    HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, (text.length() + 1) * sizeof(wchar_t));
    if (hg == NULL)
    {
        CloseClipboard();
        return;
    }

    wchar_t* pstr = static_cast<wchar_t*>(GlobalLock(hg));
    if (pstr != NULL)
    {
        wcscpy_s(pstr, text.length() + 1, text.c_str());
        GlobalUnlock(hg);
        SetClipboardData(CF_UNICODETEXT, hg);
    }

    CloseClipboard();
}

BOOL WINAPI CtrlHandler(DWORD dwCtrlType)
{
    switch (dwCtrlType)
    {
    case CTRL_C_EVENT:
    case CTRL_CLOSE_EVENT:
        PostThreadMessage(GetCurrentThreadId(), WM_QUIT, 0, 0);
        Sleep(1000);
        return TRUE;
    default:
        return FALSE;
    }
}