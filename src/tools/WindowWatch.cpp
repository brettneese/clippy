#define _WIN32_WINNT 0x0501
#define WINVER 0x0501
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>

#include <string>
#include <vector>

namespace {

const wchar_t kDumpDirectory[] = L"C:\\clippy\\dumps";
const int kHotkeyId = 1;
const UINT kWindowTextTimeoutMs = 100;

struct SnapshotContext {
    std::wstring text;
};

std::wstring Indent(int depth)
{
    return std::wstring(static_cast<size_t>(depth * 2), L' ');
}

std::wstring FormatHwnd(HWND hwnd)
{
    wchar_t buffer[32];
    _snwprintf(buffer, sizeof(buffer) / sizeof(buffer[0]),
               L"0x%08lX", static_cast<unsigned long>(
                   reinterpret_cast<ULONG_PTR>(hwnd)));
    buffer[(sizeof(buffer) / sizeof(buffer[0])) - 1] = L'\0';
    return buffer;
}

std::wstring EscapeText(const std::wstring& value)
{
    std::wstring result;
    result.reserve(value.size());
    for (size_t i = 0; i < value.size(); ++i) {
        const wchar_t character = value[i];
        switch (character) {
        case L'\\': result += L"\\\\"; break;
        case L'\"': result += L"\\\""; break;
        case L'\r': result += L"\\r"; break;
        case L'\n': result += L"\\n"; break;
        case L'\t': result += L"\\t"; break;
        default:
            if (character < 0x20) {
                wchar_t escaped[8];
                _snwprintf(escaped, sizeof(escaped) / sizeof(escaped[0]),
                           L"\\x%02X", static_cast<unsigned int>(character));
                escaped[(sizeof(escaped) / sizeof(escaped[0])) - 1] = L'\0';
                result += escaped;
            } else {
                result += character;
            }
            break;
        }
    }
    return result;
}

std::wstring GetClassNameString(HWND hwnd)
{
    wchar_t className[512];
    const int length = GetClassNameW(hwnd, className,
                                    sizeof(className) / sizeof(className[0]));
    if (length <= 0) {
        return L"";
    }
    return std::wstring(className, static_cast<size_t>(length));
}

std::wstring GetWindowTextWithoutChangingFocus(HWND hwnd)
{
    DWORD_PTR lengthResult = 0;
    if (!SendMessageTimeoutW(hwnd, WM_GETTEXTLENGTH, 0, 0,
                             SMTO_ABORTIFHUNG | SMTO_BLOCK,
                             kWindowTextTimeoutMs, &lengthResult)) {
        return L"<unavailable: timeout>";
    }

    size_t characterCount = static_cast<size_t>(lengthResult);
    if (characterCount > 32767) {
        characterCount = 32767;
    }
    std::vector<wchar_t> buffer(characterCount + 1, L'\0');
    DWORD_PTR copiedResult = 0;
    if (!SendMessageTimeoutW(hwnd, WM_GETTEXT,
                             static_cast<WPARAM>(buffer.size()),
                             reinterpret_cast<LPARAM>(&buffer[0]),
                             SMTO_ABORTIFHUNG | SMTO_BLOCK,
                             kWindowTextTimeoutMs, &copiedResult)) {
        return L"<unavailable: timeout>";
    }
    buffer[buffer.size() - 1] = L'\0';
    return std::wstring(&buffer[0]);
}

void AppendWindowRecord(HWND hwnd, int depth, SnapshotContext* context)
{
    const std::wstring windowIndent = Indent(depth);
    const std::wstring fieldIndent = Indent(depth + 1);
    DWORD processId = 0;
    const DWORD threadId = GetWindowThreadProcessId(hwnd, &processId);
    const LONG style = GetWindowLongW(hwnd, GWL_STYLE);
    const LONG extendedStyle = GetWindowLongW(hwnd, GWL_EXSTYLE);
    HWND parent = NULL;
    if ((static_cast<DWORD>(style) & WS_CHILD) != 0) {
        parent = GetParent(hwnd);
    }
    const HWND owner = GetWindow(hwnd, GW_OWNER);
    RECT rectangle;
    const BOOL hasRectangle = GetWindowRect(hwnd, &rectangle);
    const int controlId = GetDlgCtrlID(hwnd);

    wchar_t line[512];
    context->text += windowIndent + L"HWND=" + FormatHwnd(hwnd) + L"\r\n";

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"PID=%lu\r\n", static_cast<unsigned long>(processId));
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"TID=%lu\r\n", static_cast<unsigned long>(threadId));
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    context->text += fieldIndent + L"Parent=" + FormatHwnd(parent) + L"\r\n";
    context->text += fieldIndent + L"Owner=" + FormatHwnd(owner) + L"\r\n";
    context->text += fieldIndent + L"Class=\"" +
                     EscapeText(GetClassNameString(hwnd)) + L"\"\r\n";
    context->text += fieldIndent + L"Text=\"" +
                     EscapeText(GetWindowTextWithoutChangingFocus(hwnd)) +
                     L"\"\r\n";

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"ControlID=%d\r\n", controlId);
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"Visible=%d\r\n",
               IsWindowVisible(hwnd) ? 1 : 0);
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"Enabled=%d\r\n", IsWindowEnabled(hwnd) ? 1 : 0);
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    if (hasRectangle) {
        _snwprintf(line, sizeof(line) / sizeof(line[0]),
                   L"Rect=(%ld,%ld)-(%ld,%ld) Width=%ld Height=%ld\r\n",
                   rectangle.left, rectangle.top, rectangle.right,
                   rectangle.bottom, rectangle.right - rectangle.left,
                   rectangle.bottom - rectangle.top);
    } else {
        _snwprintf(line, sizeof(line) / sizeof(line[0]),
                   L"Rect=<unavailable>\r\n");
    }
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"Style=0x%08lX\r\n",
               static_cast<unsigned long>(static_cast<DWORD>(style)));
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    _snwprintf(line, sizeof(line) / sizeof(line[0]),
               L"ExStyle=0x%08lX\r\n",
               static_cast<unsigned long>(static_cast<DWORD>(extendedStyle)));
    line[(sizeof(line) / sizeof(line[0])) - 1] = L'\0';
    context->text += fieldIndent + line;

    for (HWND child = GetWindow(hwnd, GW_CHILD);
         child != NULL;
         child = GetWindow(child, GW_HWNDNEXT)) {
        AppendWindowRecord(child, depth + 1, context);
    }
}

BOOL CALLBACK AppendTopLevelWindow(HWND hwnd, LPARAM parameter)
{
    SnapshotContext* context = reinterpret_cast<SnapshotContext*>(parameter);
    AppendWindowRecord(hwnd, 0, context);
    return TRUE;
}

bool EnsureDumpDirectory()
{
    if (!CreateDirectoryW(L"C:\\clippy", NULL)) {
        const DWORD error = GetLastError();
        if (error != ERROR_ALREADY_EXISTS) {
            return false;
        }
    }
    if (!CreateDirectoryW(kDumpDirectory, NULL)) {
        const DWORD error = GetLastError();
        if (error != ERROR_ALREADY_EXISTS) {
            return false;
        }
    }
    return true;
}

unsigned int FindNextDumpNumber()
{
    std::wstring pattern = std::wstring(kDumpDirectory) + L"\\dump-*.txt";
    WIN32_FIND_DATAW findData;
    HANDLE findHandle = FindFirstFileW(pattern.c_str(), &findData);
    unsigned int maximum = 0;
    if (findHandle == INVALID_HANDLE_VALUE) {
        return 1;
    }

    do {
        unsigned int number = 0;
        wchar_t trailing = L'\0';
        if (swscanf(findData.cFileName, L"dump-%u.txt%c", &number,
                    &trailing) == 1 && number > maximum) {
            maximum = number;
        }
    } while (FindNextFileW(findHandle, &findData));
    FindClose(findHandle);
    return maximum + 1;
}

std::wstring FormatTimestamp(const SYSTEMTIME& time)
{
    wchar_t timestamp[64];
    _snwprintf(timestamp, sizeof(timestamp) / sizeof(timestamp[0]),
               L"%04u-%02u-%02u %02u:%02u:%02u.%03u local",
               time.wYear, time.wMonth, time.wDay, time.wHour,
               time.wMinute, time.wSecond, time.wMilliseconds);
    timestamp[(sizeof(timestamp) / sizeof(timestamp[0])) - 1] = L'\0';
    return timestamp;
}

bool WriteUtf8(HANDLE file, const std::wstring& text)
{
    if (text.empty()) {
        return true;
    }
    const int byteCount = WideCharToMultiByte(CP_UTF8, 0, text.data(),
                                               static_cast<int>(text.size()),
                                               NULL, 0, NULL, NULL);
    if (byteCount <= 0) {
        return false;
    }
    std::vector<char> bytes(static_cast<size_t>(byteCount));
    if (WideCharToMultiByte(CP_UTF8, 0, text.data(),
                            static_cast<int>(text.size()), &bytes[0],
                            byteCount, NULL, NULL) != byteCount) {
        return false;
    }

    size_t offset = 0;
    while (offset < bytes.size()) {
        DWORD written = 0;
        const DWORD remaining = static_cast<DWORD>(bytes.size() - offset);
        if (!WriteFile(file, &bytes[offset], remaining, &written, NULL) ||
            written == 0) {
            return false;
        }
        offset += written;
    }
    return true;
}

bool AppendIndexLine(unsigned int number,
                     const std::wstring& timestamp,
                     HWND foreground,
                     const std::wstring& foregroundClass,
                     const std::wstring& foregroundTitle)
{
    const std::wstring path = std::wstring(kDumpDirectory) + L"\\index.txt";
    HANDLE file = CreateFileW(path.c_str(), FILE_APPEND_DATA,
                              FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
                              OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    wchar_t prefix[64];
    _snwprintf(prefix, sizeof(prefix) / sizeof(prefix[0]),
               L"%03u | ", number);
    prefix[(sizeof(prefix) / sizeof(prefix[0])) - 1] = L'\0';
    const std::wstring line = std::wstring(prefix) + timestamp + L" | " +
        FormatHwnd(foreground) + L" | " + EscapeText(foregroundClass) +
        L" | " + EscapeText(foregroundTitle) + L"\r\n";
    const bool success = WriteUtf8(file, line);
    CloseHandle(file);
    return success;
}

void CaptureSnapshot()
{
    if (!EnsureDumpDirectory()) {
        return;
    }

    SYSTEMTIME localTime;
    GetLocalTime(&localTime);
    const std::wstring timestamp = FormatTimestamp(localTime);
    const HWND foreground = GetForegroundWindow();
    HWND focus = NULL;
    DWORD foregroundProcessId = 0;
    const DWORD foregroundThreadId =
        GetWindowThreadProcessId(foreground, &foregroundProcessId);
    if (foregroundThreadId != 0) {
        GUITHREADINFO guiThreadInfo;
        ZeroMemory(&guiThreadInfo, sizeof(guiThreadInfo));
        guiThreadInfo.cbSize = sizeof(guiThreadInfo);
        if (GetGUIThreadInfo(foregroundThreadId, &guiThreadInfo)) {
            focus = guiThreadInfo.hwndFocus;
        }
    }

    const std::wstring foregroundClass = GetClassNameString(foreground);
    const std::wstring foregroundTitle =
        GetWindowTextWithoutChangingFocus(foreground);

    unsigned int number = FindNextDumpNumber();
    HANDLE dumpFile = INVALID_HANDLE_VALUE;
    std::wstring dumpPath;
    for (;;) {
        wchar_t fileName[64];
        _snwprintf(fileName, sizeof(fileName) / sizeof(fileName[0]),
                   L"\\dump-%03u.txt", number);
        fileName[(sizeof(fileName) / sizeof(fileName[0])) - 1] = L'\0';
        dumpPath = std::wstring(kDumpDirectory) + fileName;
        dumpFile = CreateFileW(dumpPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                               NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
        if (dumpFile != INVALID_HANDLE_VALUE) {
            break;
        }
        if (GetLastError() != ERROR_FILE_EXISTS &&
            GetLastError() != ERROR_ALREADY_EXISTS) {
            return;
        }
        ++number;
    }

    SnapshotContext context;
    wchar_t header[128];
    _snwprintf(header, sizeof(header) / sizeof(header[0]),
               L"Dump: %03u\r\n", number);
    header[(sizeof(header) / sizeof(header[0])) - 1] = L'\0';
    context.text = header;
    context.text += L"Timestamp: " + timestamp + L"\r\n";
    context.text += L"Foreground: " + FormatHwnd(foreground) + L"\r\n";
    context.text += L"Focus: " + FormatHwnd(focus) + L"\r\n\r\n";
    EnumWindows(AppendTopLevelWindow,
                reinterpret_cast<LPARAM>(&context));

    const bool dumpWritten = WriteUtf8(dumpFile, context.text) &&
                             FlushFileBuffers(dumpFile);
    CloseHandle(dumpFile);
    if (!dumpWritten) {
        DeleteFileW(dumpPath.c_str());
        return;
    }

    AppendIndexLine(number, timestamp, foreground, foregroundClass,
                    foregroundTitle);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int)
{
    HANDLE instanceMutex = CreateMutexW(NULL, TRUE,
                                        L"Local\\WindowWatch-XP-Hotkey");
    if (instanceMutex == NULL || GetLastError() == ERROR_ALREADY_EXISTS) {
        if (instanceMutex != NULL) {
            CloseHandle(instanceMutex);
        }
        return 1;
    }

    if (!RegisterHotKey(NULL, kHotkeyId, 0, VK_F11)) {
        CloseHandle(instanceMutex);
        return 2;
    }

    MSG message;
    while (GetMessageW(&message, NULL, 0, 0) > 0) {
        if (message.message == WM_HOTKEY &&
            message.wParam == static_cast<WPARAM>(kHotkeyId)) {
            CaptureSnapshot();
        }
    }

    UnregisterHotKey(NULL, kHotkeyId);
    ReleaseMutex(instanceMutex);
    CloseHandle(instanceMutex);
    return 0;
}
