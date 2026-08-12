#define _WIN32_WINNT 0x0501
#define WINVER 0x0501
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <oleauto.h>
#include <shlwapi.h>

#include <string>
#include <vector>

namespace {

// {8D2A7B26-0D25-4D42-9E67-A62A59F9136C}
const CLSID CLSID_ClippyShim =
    {0x8d2a7b26, 0x0d25, 0x4d42,
     {0x9e, 0x67, 0xa6, 0x2a, 0x59, 0xf9, 0x13, 0x6c}};

// Office's _IDTExtensibility2 interface. Declaring the interface locally keeps
// the shim independent of Office SDK headers while preserving its COM ABI.
MIDL_INTERFACE("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")
IDTExtensibility2 : public IDispatch
{
public:
    virtual HRESULT STDMETHODCALLTYPE OnConnection(
        IDispatch* application,
        LONG connectMode,
        IDispatch* addInInstance,
        SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(
        LONG removeMode,
        SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY** custom) = 0;
};

const wchar_t kClassId[] =
    L"{8D2A7B26-0D25-4D42-9E67-A62A59F9136C}";
const wchar_t kProgId[] = L"ClippyShim.Connect";
const wchar_t kPlaceholder[] =
    L"Type your question here and then click Search.";
const wchar_t kLogPath[] = L"C:\\clippy\\ClippyShim.log";
const UINT_PTR kPollIntervalMs = 200;
const DWORD kEchoDelayMs = 150;

HINSTANCE g_module = NULL;
LONG g_objectCount = 0;
LONG g_serverLocks = 0;

HRESULT InvokeByName(IDispatch* object,
                     const wchar_t* memberName,
                     WORD flags,
                     VARIANT* arguments,
                     UINT argumentCount,
                     VARIANT* result)
{
    if (object == NULL) {
        return E_POINTER;
    }

    LPOLESTR mutableName = const_cast<LPOLESTR>(memberName);
    DISPID memberId = DISPID_UNKNOWN;
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &mutableName, 1,
                                       LOCALE_USER_DEFAULT, &memberId);
    if (FAILED(hr)) {
        return hr;
    }

    DISPID propertyPutId = DISPID_PROPERTYPUT;
    DISPPARAMS parameters;
    ZeroMemory(&parameters, sizeof(parameters));
    parameters.rgvarg = arguments;
    parameters.cArgs = argumentCount;
    if ((flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0) {
        parameters.rgdispidNamedArgs = &propertyPutId;
        parameters.cNamedArgs = 1;
    }

    VARIANT localResult;
    VariantInit(&localResult);
    VARIANT* output = result != NULL ? result : &localResult;
    if (result != NULL) {
        VariantInit(result);
    }

    EXCEPINFO exceptionInfo;
    ZeroMemory(&exceptionInfo, sizeof(exceptionInfo));
    UINT argumentError = 0;
    hr = object->Invoke(memberId, IID_NULL, LOCALE_USER_DEFAULT, flags,
                        &parameters, output, &exceptionInfo, &argumentError);
    SysFreeString(exceptionInfo.bstrSource);
    SysFreeString(exceptionInfo.bstrDescription);
    SysFreeString(exceptionInfo.bstrHelpFile);
    if (output == &localResult) {
        VariantClear(&localResult);
    }
    return hr;
}

HRESULT GetDispatchProperty(IDispatch* object,
                            const wchar_t* memberName,
                            IDispatch** value)
{
    if (value == NULL) {
        return E_POINTER;
    }
    *value = NULL;

    VARIANT result;
    HRESULT hr = InvokeByName(object, memberName, DISPATCH_PROPERTYGET,
                              NULL, 0, &result);
    if (SUCCEEDED(hr)) {
        if (result.vt == VT_DISPATCH && result.pdispVal != NULL) {
            *value = result.pdispVal;
            (*value)->AddRef();
        } else if (result.vt == VT_UNKNOWN && result.punkVal != NULL) {
            hr = result.punkVal->QueryInterface(
                IID_IDispatch, reinterpret_cast<void**>(value));
        } else {
            hr = DISP_E_TYPEMISMATCH;
        }
    }
    VariantClear(&result);
    return hr;
}

HRESULT PutString(IDispatch* object,
                  const wchar_t* memberName,
                  const std::wstring& value)
{
    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_BSTR;
    argument.bstrVal = SysAllocStringLen(
        value.data(), static_cast<UINT>(value.size()));
    if (argument.bstrVal == NULL && !value.empty()) {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = InvokeByName(object, memberName, DISPATCH_PROPERTYPUT,
                              &argument, 1, NULL);
    VariantClear(&argument);
    return hr;
}

HRESULT PutLong(IDispatch* object, const wchar_t* memberName, LONG value)
{
    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_I4;
    argument.lVal = value;
    return InvokeByName(object, memberName, DISPATCH_PROPERTYPUT,
                        &argument, 1, NULL);
}

HRESULT PutBool(IDispatch* object, const wchar_t* memberName, bool value)
{
    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_BOOL;
    argument.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return InvokeByName(object, memberName, DISPATCH_PROPERTYPUT,
                        &argument, 1, NULL);
}

std::wstring GetWindowTextString(HWND window)
{
    int length = GetWindowTextLengthW(window);
    if (length <= 0) {
        return L"";
    }
    if (length > 4096) {
        length = 4096;
    }
    std::vector<wchar_t> buffer(static_cast<size_t>(length) + 1, L'\0');
    const int copied = GetWindowTextW(window, &buffer[0], length + 1);
    if (copied <= 0) {
        return L"";
    }
    return std::wstring(&buffer[0], static_cast<size_t>(copied));
}

std::wstring Trim(const std::wstring& text)
{
    const wchar_t* whitespace = L" \t\r\n";
    const size_t first = text.find_first_not_of(whitespace);
    if (first == std::wstring::npos) {
        return L"";
    }
    const size_t last = text.find_last_not_of(whitespace);
    return text.substr(first, last - first + 1);
}

void AppendLog(const std::wstring& message)
{
    HANDLE file = CreateFileW(kLogPath, FILE_APPEND_DATA,
                              FILE_SHARE_READ | FILE_SHARE_WRITE, NULL,
                              OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        return;
    }

    SYSTEMTIME time;
    GetLocalTime(&time);
    wchar_t prefix[64];
    _snwprintf(prefix, sizeof(prefix) / sizeof(prefix[0]),
               L"%04u-%02u-%02u %02u:%02u:%02u ",
               time.wYear, time.wMonth, time.wDay,
               time.wHour, time.wMinute, time.wSecond);
    prefix[(sizeof(prefix) / sizeof(prefix[0])) - 1] = L'\0';
    const std::wstring line = std::wstring(prefix) + message + L"\r\n";

    const int byteCount = WideCharToMultiByte(
        CP_UTF8, 0, line.data(), static_cast<int>(line.size()),
        NULL, 0, NULL, NULL);
    if (byteCount > 0) {
        std::vector<char> bytes(static_cast<size_t>(byteCount));
        if (WideCharToMultiByte(CP_UTF8, 0, line.data(),
                                static_cast<int>(line.size()), &bytes[0],
                                byteCount, NULL, NULL) == byteCount) {
            DWORD written = 0;
            WriteFile(file, &bytes[0], static_cast<DWORD>(bytes.size()),
                      &written, NULL);
        }
    }
    CloseHandle(file);
}

bool IsClass(HWND window, const wchar_t* expected)
{
    wchar_t className[128];
    const int length = GetClassNameW(
        window, className, sizeof(className) / sizeof(className[0]));
    return length > 0 && lstrcmpW(className, expected) == 0;
}

struct BalloonWindows {
    HWND shell;
    HWND content;
    HWND editor;

    BalloonWindows() : shell(NULL), content(NULL), editor(NULL) {}
};

BOOL CALLBACK FindBalloonWindow(HWND shell, LPARAM parameter)
{
    BalloonWindows* result = reinterpret_cast<BalloonWindows*>(parameter);
    if (!IsWindowVisible(shell) || !IsClass(shell, L"MSOBALLOON")) {
        return TRUE;
    }

    for (HWND content = FindWindowExW(shell, NULL, L"MsoBalloonChild", NULL);
         content != NULL;
         content = FindWindowExW(shell, content, L"MsoBalloonChild", NULL)) {
        DWORD processId = 0;
        GetWindowThreadProcessId(content, &processId);
        if (processId != GetCurrentProcessId()) {
            continue;
        }
        HWND editor = GetDlgItem(content, 6);
        if (editor != NULL && IsClass(editor, L"RichEdit20W")) {
            result->shell = shell;
            result->content = content;
            result->editor = editor;
            return FALSE;
        }
    }
    return TRUE;
}

class ClippyAddIn;
ClippyAddIn* g_activeAddIn = NULL;

class ClippyAddIn : public IDTExtensibility2
{
public:
    ClippyAddIn()
        : referenceCount_(1), application_(NULL), timer_(0), shell_(NULL),
          content_(NULL), editor_(NULL), oldContentProc_(NULL),
          oldEditorProc_(NULL), echoPending_(false), echoDueTick_(0)
    {
        InterlockedIncrement(&g_objectCount);
    }

    virtual ~ClippyAddIn()
    {
        Stop();
        InterlockedDecrement(&g_objectCount);
    }

    STDMETHODIMP QueryInterface(REFIID interfaceId, void** object)
    {
        if (object == NULL) {
            return E_POINTER;
        }
        *object = NULL;
        if (interfaceId == IID_IUnknown ||
            interfaceId == IID_IDispatch ||
            interfaceId == __uuidof(IDTExtensibility2)) {
            *object = static_cast<IDTExtensibility2*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef()
    {
        return static_cast<ULONG>(InterlockedIncrement(&referenceCount_));
    }

    STDMETHODIMP_(ULONG) Release()
    {
        const LONG count = InterlockedDecrement(&referenceCount_);
        if (count == 0) {
            delete this;
            return 0;
        }
        return static_cast<ULONG>(count);
    }

    STDMETHODIMP GetTypeInfoCount(UINT* count)
    {
        if (count == NULL) {
            return E_POINTER;
        }
        *count = 0;
        return S_OK;
    }

    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo** typeInfo)
    {
        if (typeInfo != NULL) {
            *typeInfo = NULL;
        }
        return E_NOTIMPL;
    }

    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*)
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*,
                        VARIANT*, EXCEPINFO*, UINT*)
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP OnConnection(IDispatch* application,
                              LONG,
                              IDispatch*,
                              SAFEARRAY**)
    {
        Stop();
        application_ = application;
        if (application_ != NULL) {
            application_->AddRef();
        }
        g_activeAddIn = this;
        timer_ = SetTimer(NULL, 0, kPollIntervalMs, TimerCallback);
        AppendLog(timer_ != 0 ? L"ADDIN connected" :
                                L"ERROR SetTimer failed");
        return timer_ != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    STDMETHODIMP OnDisconnection(LONG, SAFEARRAY**)
    {
        AppendLog(L"ADDIN disconnected");
        Stop();
        return S_OK;
    }

    STDMETHODIMP OnAddInsUpdate(SAFEARRAY**) { return S_OK; }
    STDMETHODIMP OnStartupComplete(SAFEARRAY**) { return S_OK; }
    STDMETHODIMP OnBeginShutdown(SAFEARRAY**) { return S_OK; }

    static void CALLBACK TimerCallback(HWND, UINT, UINT_PTR, DWORD)
    {
        if (g_activeAddIn != NULL) {
            g_activeAddIn->Poll();
        }
    }

    static LRESULT CALLBACK ContentWindowProc(HWND window,
                                               UINT message,
                                               WPARAM wParam,
                                               LPARAM lParam)
    {
        ClippyAddIn* self = g_activeAddIn;
        if (self == NULL || window != self->content_) {
            return DefWindowProcW(window, message, wParam, lParam);
        }

        WNDPROC original = self->oldContentProc_;
        if (message == WM_COMMAND &&
            LOWORD(wParam) == 8 && HIWORD(wParam) == BN_CLICKED &&
            reinterpret_cast<HWND>(lParam) == self->editorSearchButton()) {
            if (self->CaptureQuery(L"Search button")) {
                return 0;
            }
        }

        const LRESULT result = CallWindowProcW(
            original, window, message, wParam, lParam);
        if (message == WM_NCDESTROY) {
            self->content_ = NULL;
            self->oldContentProc_ = NULL;
        }
        return result;
    }

    static LRESULT CALLBACK EditorWindowProc(HWND window,
                                              UINT message,
                                              WPARAM wParam,
                                              LPARAM lParam)
    {
        ClippyAddIn* self = g_activeAddIn;
        if (self == NULL || window != self->editor_) {
            return DefWindowProcW(window, message, wParam, lParam);
        }

        WNDPROC original = self->oldEditorProc_;
        if (message == WM_KEYDOWN && wParam == VK_RETURN &&
            (GetKeyState(VK_SHIFT) & 0x8000) == 0 &&
            (GetKeyState(VK_CONTROL) & 0x8000) == 0) {
            if (self->CaptureQuery(L"Enter key")) {
                return 0;
            }
        }

        const LRESULT result = CallWindowProcW(
            original, window, message, wParam, lParam);
        if (message == WM_NCDESTROY) {
            self->editor_ = NULL;
            self->oldEditorProc_ = NULL;
        }
        return result;
    }

private:
    HWND editorSearchButton() const
    {
        return content_ != NULL ? GetDlgItem(content_, 8) : NULL;
    }

    void Stop()
    {
        if (timer_ != 0) {
            KillTimer(NULL, timer_);
            timer_ = 0;
        }
        DetachHooks();
        if (g_activeAddIn == this) {
            g_activeAddIn = NULL;
        }
        if (application_ != NULL) {
            application_->Release();
            application_ = NULL;
        }
        echoPending_ = false;
        pendingQuery_.clear();
    }

    void Poll()
    {
        if (echoPending_) {
            const DWORD now = GetTickCount();
            if (static_cast<LONG>(now - echoDueTick_) >= 0) {
                echoPending_ = false;
                const std::wstring query = pendingQuery_;
                pendingQuery_.clear();
                ShowEcho(query);
            }
            return;
        }

        if (content_ != NULL && editor_ != NULL &&
            IsWindow(content_) && IsWindow(editor_)) {
            return;
        }

        DetachHooks();
        BalloonWindows windows;
        EnumWindows(FindBalloonWindow, reinterpret_cast<LPARAM>(&windows));
        if (windows.content == NULL || windows.editor == NULL) {
            return;
        }

        shell_ = windows.shell;
        content_ = windows.content;
        editor_ = windows.editor;
        SetLastError(ERROR_SUCCESS);
        oldContentProc_ = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(
            content_, GWLP_WNDPROC,
            reinterpret_cast<LONG_PTR>(ContentWindowProc)));
        if (oldContentProc_ == NULL && GetLastError() != ERROR_SUCCESS) {
            AppendLog(L"ERROR subclassing MsoBalloonChild");
            content_ = NULL;
            editor_ = NULL;
            shell_ = NULL;
            return;
        }

        SetLastError(ERROR_SUCCESS);
        oldEditorProc_ = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(
            editor_, GWLP_WNDPROC,
            reinterpret_cast<LONG_PTR>(EditorWindowProc)));
        if (oldEditorProc_ == NULL && GetLastError() != ERROR_SUCCESS) {
            SetWindowLongPtrW(content_, GWLP_WNDPROC,
                              reinterpret_cast<LONG_PTR>(oldContentProc_));
            oldContentProc_ = NULL;
            content_ = NULL;
            editor_ = NULL;
            shell_ = NULL;
            AppendLog(L"ERROR subclassing RichEdit20W");
            return;
        }
        AppendLog(L"HOOK attached to Office Assistant query balloon");
    }

    void DetachHooks()
    {
        if (editor_ != NULL && IsWindow(editor_) && oldEditorProc_ != NULL) {
            SetWindowLongPtrW(editor_, GWLP_WNDPROC,
                              reinterpret_cast<LONG_PTR>(oldEditorProc_));
        }
        if (content_ != NULL && IsWindow(content_) && oldContentProc_ != NULL) {
            SetWindowLongPtrW(content_, GWLP_WNDPROC,
                              reinterpret_cast<LONG_PTR>(oldContentProc_));
        }
        oldEditorProc_ = NULL;
        oldContentProc_ = NULL;
        editor_ = NULL;
        content_ = NULL;
        shell_ = NULL;
    }

    bool CaptureQuery(const wchar_t* source)
    {
        if (echoPending_ || editor_ == NULL) {
            return echoPending_;
        }
        const std::wstring query = Trim(GetWindowTextString(editor_));
        if (query.empty() || query == kPlaceholder) {
            return false;
        }

        AppendLog(std::wstring(L"QUERY source=") + source + L" text=" + query);
        pendingQuery_ = query;
        echoPending_ = true;
        echoDueTick_ = GetTickCount() + kEchoDelayMs;

        // This is both immediate feedback and the fallback if the Office
        // Assistant refuses to replace its own built-in search balloon.
        const std::wstring feedback = L"You said: " + query;
        SetWindowTextW(editor_, feedback.c_str());
        return true;
    }

    void ShowEcho(const std::wstring& query)
    {
        HWND previousShell = shell_;
        DetachHooks();
        if (previousShell != NULL && IsWindow(previousShell)) {
            ShowWindow(previousShell, SW_HIDE);
        }

        IDispatch* assistant = NULL;
        IDispatch* balloon = NULL;
        HRESULT hr = GetDispatchProperty(application_, L"Assistant", &assistant);
        if (SUCCEEDED(hr)) {
            PutBool(assistant, L"Visible", true);
            VARIANT result;
            hr = InvokeByName(assistant, L"NewBalloon",
                              DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                              NULL, 0, &result);
            if (SUCCEEDED(hr)) {
                if (result.vt == VT_DISPATCH && result.pdispVal != NULL) {
                    balloon = result.pdispVal;
                    balloon->AddRef();
                } else if (result.vt == VT_UNKNOWN && result.punkVal != NULL) {
                    hr = result.punkVal->QueryInterface(
                        IID_IDispatch, reinterpret_cast<void**>(&balloon));
                } else {
                    hr = DISP_E_TYPEMISMATCH;
                }
            }
            VariantClear(&result);
        }

        if (SUCCEEDED(hr) && balloon != NULL) {
            hr = PutString(balloon, L"Heading", L"Clippy heard:");
            if (SUCCEEDED(hr)) {
                hr = PutString(balloon, L"Text", query);
            }
            if (SUCCEEDED(hr)) {
                hr = PutLong(balloon, L"Button", 1);  // msoButtonSetOK
            }
            if (SUCCEEDED(hr)) {
                hr = InvokeByName(balloon, L"Show", DISPATCH_METHOD,
                                  NULL, 0, NULL);
            }
        }

        if (balloon != NULL) {
            balloon->Release();
        }
        if (assistant != NULL) {
            assistant->Release();
        }

        if (SUCCEEDED(hr)) {
            AppendLog(L"ECHO shown with Assistant.NewBalloon");
        } else {
            wchar_t error[64];
            _snwprintf(error, sizeof(error) / sizeof(error[0]),
                       L"ERROR echo failed hr=0x%08lX",
                       static_cast<unsigned long>(hr));
            error[(sizeof(error) / sizeof(error[0])) - 1] = L'\0';
            AppendLog(error);
            if (previousShell != NULL && IsWindow(previousShell)) {
                ShowWindow(previousShell, SW_SHOWNA);
            }
        }
    }

    LONG referenceCount_;
    IDispatch* application_;
    UINT_PTR timer_;
    HWND shell_;
    HWND content_;
    HWND editor_;
    WNDPROC oldContentProc_;
    WNDPROC oldEditorProc_;
    bool echoPending_;
    DWORD echoDueTick_;
    std::wstring pendingQuery_;
};

class ClippyClassFactory : public IClassFactory
{
public:
    ClippyClassFactory() : referenceCount_(1) {}

    STDMETHODIMP QueryInterface(REFIID interfaceId, void** object)
    {
        if (object == NULL) {
            return E_POINTER;
        }
        *object = NULL;
        if (interfaceId == IID_IUnknown || interfaceId == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef()
    {
        return static_cast<ULONG>(InterlockedIncrement(&referenceCount_));
    }

    STDMETHODIMP_(ULONG) Release()
    {
        const LONG count = InterlockedDecrement(&referenceCount_);
        if (count == 0) {
            delete this;
            return 0;
        }
        return static_cast<ULONG>(count);
    }

    STDMETHODIMP CreateInstance(IUnknown* outer,
                                REFIID interfaceId,
                                void** object)
    {
        if (outer != NULL) {
            return CLASS_E_NOAGGREGATION;
        }
        ClippyAddIn* addIn = new ClippyAddIn();
        if (addIn == NULL) {
            return E_OUTOFMEMORY;
        }
        const HRESULT hr = addIn->QueryInterface(interfaceId, object);
        addIn->Release();
        return hr;
    }

    STDMETHODIMP LockServer(BOOL lock)
    {
        if (lock) {
            InterlockedIncrement(&g_serverLocks);
        } else {
            InterlockedDecrement(&g_serverLocks);
        }
        return S_OK;
    }

private:
    LONG referenceCount_;
};

LONG SetStringValue(HKEY root,
                    const std::wstring& path,
                    const wchar_t* name,
                    const std::wstring& value)
{
    HKEY key = NULL;
    LONG status = RegCreateKeyExW(root, path.c_str(), 0, NULL, 0,
                                  KEY_SET_VALUE, NULL, &key, NULL);
    if (status == ERROR_SUCCESS) {
        status = RegSetValueExW(
            key, name, 0, REG_SZ,
            reinterpret_cast<const BYTE*>(value.c_str()),
            static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));
        RegCloseKey(key);
    }
    return status;
}

LONG SetDwordValue(HKEY root,
                   const std::wstring& path,
                   const wchar_t* name,
                   DWORD value)
{
    HKEY key = NULL;
    LONG status = RegCreateKeyExW(root, path.c_str(), 0, NULL, 0,
                                  KEY_SET_VALUE, NULL, &key, NULL);
    if (status == ERROR_SUCCESS) {
        status = RegSetValueExW(key, name, 0, REG_DWORD,
                                reinterpret_cast<const BYTE*>(&value),
                                sizeof(value));
        RegCloseKey(key);
    }
    return status;
}

}  // namespace

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH) {
        g_module = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return g_objectCount == 0 && g_serverLocks == 0 ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID classId,
                                                REFIID interfaceId,
                                                void** object)
{
    if (classId != CLSID_ClippyShim) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    ClippyClassFactory* factory = new ClippyClassFactory();
    if (factory == NULL) {
        return E_OUTOFMEMORY;
    }
    const HRESULT hr = factory->QueryInterface(interfaceId, object);
    factory->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllRegisterServer()
{
    wchar_t modulePath[MAX_PATH];
    const DWORD length = GetModuleFileNameW(
        g_module, modulePath, sizeof(modulePath) / sizeof(modulePath[0]));
    if (length == 0 || length >= sizeof(modulePath) / sizeof(modulePath[0])) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::wstring classPath =
        std::wstring(L"Software\\Classes\\CLSID\\") + kClassId;
    const std::wstring progIdPath =
        std::wstring(L"Software\\Classes\\") + kProgId;
    const std::wstring addInPath =
        std::wstring(L"Software\\Microsoft\\Office\\Word\\Addins\\") +
        kProgId;

    LONG status = SetStringValue(HKEY_CURRENT_USER, classPath, NULL,
                                 L"ClippyShim Word Add-in");
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER,
                                classPath + L"\\InprocServer32", NULL,
                                modulePath);
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER,
                                classPath + L"\\InprocServer32",
                                L"ThreadingModel", L"Apartment");
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER,
                                classPath + L"\\ProgID", NULL, kProgId);
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER, progIdPath, NULL,
                                L"ClippyShim Word Add-in");
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER,
                                progIdPath + L"\\CLSID", NULL, kClassId);
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(HKEY_CURRENT_USER, addInPath,
                                L"FriendlyName", L"Clippy Possession Shim");
    }
    if (status == ERROR_SUCCESS) {
        status = SetStringValue(
            HKEY_CURRENT_USER, addInPath, L"Description",
            L"Intercepts Office Assistant questions for Clippy Possession.");
    }
    if (status == ERROR_SUCCESS) {
        status = SetDwordValue(HKEY_CURRENT_USER, addInPath,
                               L"LoadBehavior", 3);
    }
    if (status == ERROR_SUCCESS) {
        status = SetDwordValue(HKEY_CURRENT_USER, addInPath,
                               L"CommandLineSafe", 0);
    }
    return HRESULT_FROM_WIN32(status);
}

extern "C" HRESULT __stdcall DllUnregisterServer()
{
    const std::wstring classPath =
        std::wstring(L"Software\\Classes\\CLSID\\") + kClassId;
    const std::wstring progIdPath =
        std::wstring(L"Software\\Classes\\") + kProgId;
    const std::wstring addInPath =
        std::wstring(L"Software\\Microsoft\\Office\\Word\\Addins\\") +
        kProgId;
    SHDeleteKeyW(HKEY_CURRENT_USER, addInPath.c_str());
    SHDeleteKeyW(HKEY_CURRENT_USER, progIdPath.c_str());
    SHDeleteKeyW(HKEY_CURRENT_USER, classPath.c_str());
    return S_OK;
}
