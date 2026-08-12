#define _WIN32_WINNT 0x0501
#define WINVER 0x0501
#define _CRT_SECURE_NO_WARNINGS

#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <oleauto.h>
#include <shlwapi.h>

#include <cstdio>
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
const char kDefaultServerHost[] = "10.0.2.2";
const unsigned short kDefaultServerPort = 3210;
const int kConnectTimeoutMs = 3000;
const int kSendTimeoutMs = 3000;
const int kReceiveTimeoutMs = 60000;
const size_t kMaximumResponseBytes = 64 * 1024;
// Office 10 MsoAnimationType values used through late-bound automation.
const LONG kMsoAnimationIdle = 1;
const LONG kMsoAnimationThinking = 24;

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
    if (result != NULL) {
        VariantInit(result);
    }
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

HRESULT GetIndexedDispatch(IDispatch* collection,
                           LONG index,
                           IDispatch** value)
{
    if (value == NULL) {
        return E_POINTER;
    }
    *value = NULL;

    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_I4;
    argument.lVal = index;

    VARIANT result;
    HRESULT hr = InvokeByName(
        collection, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        &argument, 1, &result);
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

struct BalloonMarkup {
    std::wstring text;
    LONG listType;
    std::vector<std::wstring> labels;

    BalloonMarkup() : listType(0) {}
};

void AppendBalloonLiteral(std::wstring* output,
                          const std::wstring& text)
{
    for (size_t index = 0; index < text.size(); ++index) {
        const wchar_t character = text[index];
        // Office uses brace-delimited directives for color, underlining, and
        // local BMP/WMF references. Full-width braces preserve readability
        // without allowing model output to inject one of those directives.
        if (character == L'{') {
            output->push_back(static_cast<wchar_t>(0xff5b));
        } else if (character == L'}') {
            output->push_back(static_cast<wchar_t>(0xff5d));
        } else {
            output->push_back(character);
        }
    }
}

bool IsMarkdownEscapable(wchar_t character)
{
    return character == L'\\' || character == L'`' ||
           character == L'*' || character == L'_' ||
           character == L'[' || character == L']' ||
           character == L'(' || character == L')' ||
           character == L'{' || character == L'}' ||
           character == L'#';
}

std::wstring FormatInlineMarkdown(const std::wstring& text)
{
    std::wstring output;
    size_t position = 0;
    while (position < text.size()) {
        if (text[position] == L'\\' && position + 1 < text.size() &&
            IsMarkdownEscapable(text[position + 1])) {
            AppendBalloonLiteral(&output, text.substr(position + 1, 1));
            position += 2;
            continue;
        }

        if (text[position] == L'`') {
            const size_t close = text.find(L'`', position + 1);
            if (close != std::wstring::npos && close > position + 1) {
                output += L"{cf 4}`";
                AppendBalloonLiteral(
                    &output, text.substr(position + 1, close - position - 1));
                output += L"`{cf 0}";
                position = close + 1;
                continue;
            }
        }

        const bool strong =
            position + 1 < text.size() &&
            ((text[position] == L'*' && text[position + 1] == L'*') ||
             (text[position] == L'_' && text[position + 1] == L'_'));
        if (strong) {
            const std::wstring delimiter = text.substr(position, 2);
            const size_t close = text.find(delimiter, position + 2);
            if (close != std::wstring::npos && close > position + 2) {
                output += L"{ul}{cf 4}";
                AppendBalloonLiteral(
                    &output, text.substr(position + 2, close - position - 2));
                output += L"{cf 0}{ul 0}";
                position = close + 2;
                continue;
            }
        }

        if (text[position] == L'*' || text[position] == L'_') {
            const wchar_t delimiter = text[position];
            const size_t close = text.find(delimiter, position + 1);
            if (close != std::wstring::npos && close > position + 1) {
                output += L"{ul}";
                AppendBalloonLiteral(
                    &output, text.substr(position + 1, close - position - 1));
                output += L"{ul 0}";
                position = close + 1;
                continue;
            }
        }

        if (text[position] == L'[') {
            const size_t labelEnd = text.find(L']', position + 1);
            if (labelEnd != std::wstring::npos &&
                labelEnd + 1 < text.size() && text[labelEnd + 1] == L'(') {
                const size_t urlEnd = text.find(L')', labelEnd + 2);
                if (urlEnd != std::wstring::npos && labelEnd > position + 1 &&
                    urlEnd > labelEnd + 2) {
                    output += L"{ul}{cf 4}";
                    AppendBalloonLiteral(
                        &output,
                        text.substr(position + 1, labelEnd - position - 1));
                    output += L"{cf 0}{ul 0} (";
                    AppendBalloonLiteral(
                        &output,
                        text.substr(labelEnd + 2, urlEnd - labelEnd - 2));
                    output += L")";
                    position = urlEnd + 1;
                    continue;
                }
            }
        }

        AppendBalloonLiteral(&output, text.substr(position, 1));
        ++position;
    }
    return output;
}

std::wstring FormatMarkdownLine(const std::wstring& line)
{
    size_t headingLength = 0;
    while (headingLength < line.size() && headingLength < 3 &&
           line[headingLength] == L'#') {
        ++headingLength;
    }
    if (headingLength > 0 && headingLength < line.size() &&
        line[headingLength] == L' ') {
        return L"{ul}{cf 4}" +
            FormatInlineMarkdown(line.substr(headingLength + 1)) +
            L"{cf 0}{ul 0}";
    }
    return FormatInlineMarkdown(line);
}

void SplitLines(const std::wstring& text,
                std::vector<std::wstring>* lines)
{
    size_t start = 0;
    while (start <= text.size()) {
        const size_t end = text.find(L'\n', start);
        std::wstring line = text.substr(
            start, end == std::wstring::npos ? std::wstring::npos : end - start);
        if (!line.empty() && line[line.size() - 1] == L'\r') {
            line.erase(line.size() - 1);
        }
        lines->push_back(line);
        if (end == std::wstring::npos) {
            break;
        }
        start = end + 1;
    }
}

bool ParseListItem(const std::wstring& line,
                   LONG* listType,
                   std::wstring* item)
{
    size_t position = 0;
    while (position < line.size() &&
           (line[position] == L' ' || line[position] == L'\t')) {
        ++position;
    }
    if (position + 1 < line.size() &&
        (line[position] == L'-' || line[position] == L'*' ||
         line[position] == L'+') && line[position + 1] == L' ') {
        *listType = 1;  // msoBalloonTypeBullets
        *item = Trim(line.substr(position + 2));
        return !item->empty();
    }

    const size_t numberStart = position;
    while (position < line.size() &&
           line[position] >= L'0' && line[position] <= L'9') {
        ++position;
    }
    if (position > numberStart && position + 1 < line.size() &&
        line[position] == L'.' && line[position + 1] == L' ') {
        *listType = 2;  // msoBalloonTypeNumbers
        *item = Trim(line.substr(position + 2));
        return !item->empty();
    }
    return false;
}

std::wstring FormatMarkdownLines(const std::vector<std::wstring>& lines,
                                 size_t count)
{
    std::wstring output;
    bool inCodeBlock = false;
    for (size_t index = 0; index < count; ++index) {
        const std::wstring trimmed = Trim(lines[index]);
        if (trimmed.size() >= 3 && trimmed.substr(0, 3) == L"```") {
            inCodeBlock = !inCodeBlock;
            continue;
        }
        if (!output.empty()) {
            output += L"\r\n";
        }
        if (inCodeBlock && !lines[index].empty()) {
            output += L"{cf 4}| ";
            AppendBalloonLiteral(&output, lines[index]);
            output += L"{cf 0}";
        } else {
            output += FormatMarkdownLine(lines[index]);
        }
    }
    while (output.size() >= 2 &&
           output.substr(output.size() - 2) == L"\r\n") {
        output.erase(output.size() - 2);
    }
    return output;
}

BalloonMarkup FormatBalloonMarkup(const std::wstring& markdown)
{
    BalloonMarkup result;
    std::vector<std::wstring> lines;
    SplitLines(markdown, &lines);
    while (!lines.empty() && Trim(lines.back()).empty()) {
        lines.pop_back();
    }

    size_t listStart = lines.size();
    LONG listType = 0;
    for (size_t index = 0; index < lines.size(); ++index) {
        LONG candidateType = 0;
        std::wstring candidate;
        if (ParseListItem(lines[index], &candidateType, &candidate)) {
            listStart = index;
            listType = candidateType;
            break;
        }
    }

    bool nativeList = listStart < lines.size();
    std::vector<std::wstring> listItems;
    if (nativeList) {
        for (size_t index = listStart; index < lines.size(); ++index) {
            LONG candidateType = 0;
            std::wstring candidate;
            if (!ParseListItem(lines[index], &candidateType, &candidate) ||
                candidateType != listType || listItems.size() >= 5) {
                nativeList = false;
                listItems.clear();
                break;
            }
            listItems.push_back(FormatInlineMarkdown(candidate));
        }
    }

    result.text = FormatMarkdownLines(
        lines, nativeList ? listStart : lines.size());
    if (nativeList) {
        result.listType = listType;
        result.labels = listItems;
    }
    return result;
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

std::string WideToUtf8(const std::wstring& text)
{
    if (text.empty()) {
        return "";
    }
    const int length = WideCharToMultiByte(
        CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
        NULL, 0, NULL, NULL);
    if (length <= 0) {
        return "";
    }
    std::vector<char> buffer(static_cast<size_t>(length));
    if (WideCharToMultiByte(
            CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
            &buffer[0], length, NULL, NULL) != length) {
        return "";
    }
    return std::string(&buffer[0], buffer.size());
}

std::wstring Utf8ToWide(const std::string& text)
{
    if (text.empty()) {
        return L"";
    }
    const int length = MultiByteToWideChar(
        CP_UTF8, 0, text.data(), static_cast<int>(text.size()), NULL, 0);
    if (length <= 0) {
        return L"";
    }
    std::vector<wchar_t> buffer(static_cast<size_t>(length));
    if (MultiByteToWideChar(
            CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
            &buffer[0], length) != length) {
        return L"";
    }
    return std::wstring(&buffer[0], buffer.size());
}

std::string EscapeJsonString(const std::string& text)
{
    std::string escaped;
    escaped.reserve(text.size() + 16);
    static const char hex[] = "0123456789abcdef";
    for (size_t index = 0; index < text.size(); ++index) {
        const unsigned char character =
            static_cast<unsigned char>(text[index]);
        switch (character) {
            case '"': escaped += "\\\""; break;
            case '\\': escaped += "\\\\"; break;
            case '\b': escaped += "\\b"; break;
            case '\f': escaped += "\\f"; break;
            case '\n': escaped += "\\n"; break;
            case '\r': escaped += "\\r"; break;
            case '\t': escaped += "\\t"; break;
            default:
                if (character < 0x20) {
                    escaped += "\\u00";
                    escaped += hex[(character >> 4) & 0x0f];
                    escaped += hex[character & 0x0f];
                } else {
                    escaped += static_cast<char>(character);
                }
                break;
        }
    }
    return escaped;
}

int HexDigit(char character)
{
    if (character >= '0' && character <= '9') {
        return character - '0';
    }
    if (character >= 'a' && character <= 'f') {
        return character - 'a' + 10;
    }
    if (character >= 'A' && character <= 'F') {
        return character - 'A' + 10;
    }
    return -1;
}

void AppendUtf8CodePoint(std::string* text, unsigned long codePoint)
{
    if (codePoint <= 0x7f) {
        text->push_back(static_cast<char>(codePoint));
    } else if (codePoint <= 0x7ff) {
        text->push_back(static_cast<char>(0xc0 | (codePoint >> 6)));
        text->push_back(static_cast<char>(0x80 | (codePoint & 0x3f)));
    } else if (codePoint <= 0xffff) {
        text->push_back(static_cast<char>(0xe0 | (codePoint >> 12)));
        text->push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3f)));
        text->push_back(static_cast<char>(0x80 | (codePoint & 0x3f)));
    } else {
        text->push_back(static_cast<char>(0xf0 | (codePoint >> 18)));
        text->push_back(static_cast<char>(0x80 | ((codePoint >> 12) & 0x3f)));
        text->push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3f)));
        text->push_back(static_cast<char>(0x80 | (codePoint & 0x3f)));
    }
}

bool ParseHexCodeUnit(const std::string& text,
                      size_t offset,
                      unsigned long* value)
{
    if (offset + 4 > text.size()) {
        return false;
    }
    unsigned long result = 0;
    for (size_t index = 0; index < 4; ++index) {
        const int digit = HexDigit(text[offset + index]);
        if (digit < 0) {
            return false;
        }
        result = (result << 4) | static_cast<unsigned long>(digit);
    }
    *value = result;
    return true;
}

bool ExtractJsonStringProperty(const std::string& json,
                               const char* propertyName,
                               std::wstring* result)
{
    const std::string quotedProperty =
        std::string("\"") + propertyName + "\"";
    const size_t property = json.find(quotedProperty);
    if (property == std::string::npos) {
        return false;
    }
    size_t position = json.find(':', property + quotedProperty.size());
    if (position == std::string::npos) {
        return false;
    }
    ++position;
    while (position < json.size() &&
           (json[position] == ' ' || json[position] == '\t' ||
            json[position] == '\r' || json[position] == '\n')) {
        ++position;
    }
    if (position >= json.size() || json[position] != '"') {
        return false;
    }
    ++position;

    std::string decoded;
    while (position < json.size()) {
        const char character = json[position++];
        if (character == '"') {
            *result = Utf8ToWide(decoded);
            return !decoded.empty() && !result->empty();
        }
        if (character != '\\') {
            decoded.push_back(character);
            continue;
        }
        if (position >= json.size()) {
            return false;
        }
        const char escape = json[position++];
        switch (escape) {
            case '"': decoded.push_back('"'); break;
            case '\\': decoded.push_back('\\'); break;
            case '/': decoded.push_back('/'); break;
            case 'b': decoded.push_back('\b'); break;
            case 'f': decoded.push_back('\f'); break;
            case 'n': decoded.push_back('\n'); break;
            case 'r': decoded.push_back('\r'); break;
            case 't': decoded.push_back('\t'); break;
            case 'u': {
                unsigned long codePoint = 0;
                if (!ParseHexCodeUnit(json, position, &codePoint)) {
                    return false;
                }
                position += 4;
                if (codePoint >= 0xd800 && codePoint <= 0xdbff &&
                    position + 6 <= json.size() &&
                    json[position] == '\\' && json[position + 1] == 'u') {
                    unsigned long lowSurrogate = 0;
                    if (ParseHexCodeUnit(json, position + 2, &lowSurrogate) &&
                        lowSurrogate >= 0xdc00 && lowSurrogate <= 0xdfff) {
                        codePoint = 0x10000 +
                            ((codePoint - 0xd800) << 10) +
                            (lowSurrogate - 0xdc00);
                        position += 6;
                    }
                }
                AppendUtf8CodePoint(&decoded, codePoint);
                break;
            }
            default:
                return false;
        }
    }
    return false;
}

std::wstring SocketError(const wchar_t* operation, int error)
{
    wchar_t message[128];
    _snwprintf(message, sizeof(message) / sizeof(message[0]),
               L"%s failed (Winsock error %d)", operation, error);
    message[(sizeof(message) / sizeof(message[0])) - 1] = L'\0';
    return message;
}

bool SendAll(SOCKET socketHandle, const std::string& data, std::wstring* error)
{
    size_t offset = 0;
    while (offset < data.size()) {
        const int sent = send(
            socketHandle, data.data() + offset,
            static_cast<int>(data.size() - offset), 0);
        if (sent == SOCKET_ERROR) {
            *error = SocketError(L"send", WSAGetLastError());
            return false;
        }
        if (sent == 0) {
            *error = L"send returned without writing data";
            return false;
        }
        offset += static_cast<size_t>(sent);
    }
    return true;
}

bool ConnectToServer(const std::string& host,
                     unsigned short port,
                     SOCKET* connectedSocket,
                     std::wstring* error)
{
    unsigned long address = inet_addr(host.c_str());
    if (address == INADDR_NONE) {
        hostent* entry = gethostbyname(host.c_str());
        if (entry == NULL || entry->h_addrtype != AF_INET ||
            entry->h_length != sizeof(address)) {
            *error = SocketError(L"name lookup", WSAGetLastError());
            return false;
        }
        CopyMemory(&address, entry->h_addr_list[0], sizeof(address));
    }

    SOCKET socketHandle = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (socketHandle == INVALID_SOCKET) {
        *error = SocketError(L"socket", WSAGetLastError());
        return false;
    }

    u_long nonBlocking = 1;
    ioctlsocket(socketHandle, FIONBIO, &nonBlocking);
    sockaddr_in serverAddress;
    ZeroMemory(&serverAddress, sizeof(serverAddress));
    serverAddress.sin_family = AF_INET;
    serverAddress.sin_port = htons(port);
    serverAddress.sin_addr.s_addr = address;

    int status = connect(
        socketHandle, reinterpret_cast<sockaddr*>(&serverAddress),
        sizeof(serverAddress));
    if (status == SOCKET_ERROR) {
        const int connectError = WSAGetLastError();
        if (connectError != WSAEWOULDBLOCK &&
            connectError != WSAEINPROGRESS) {
            *error = SocketError(L"connect", connectError);
            closesocket(socketHandle);
            return false;
        }

        fd_set writeSet;
        fd_set errorSet;
        writeSet.fd_count = 1;
        writeSet.fd_array[0] = socketHandle;
        errorSet.fd_count = 1;
        errorSet.fd_array[0] = socketHandle;
        timeval timeout;
        timeout.tv_sec = kConnectTimeoutMs / 1000;
        timeout.tv_usec = (kConnectTimeoutMs % 1000) * 1000;
        status = select(0, NULL, &writeSet, &errorSet, &timeout);
        if (status <= 0) {
            *error = status == 0 ? L"connect timed out" :
                SocketError(L"select", WSAGetLastError());
            closesocket(socketHandle);
            return false;
        }

        int socketError = 0;
        int errorLength = sizeof(socketError);
        if (getsockopt(socketHandle, SOL_SOCKET, SO_ERROR,
                       reinterpret_cast<char*>(&socketError),
                       &errorLength) == SOCKET_ERROR || socketError != 0) {
            if (socketError == 0) {
                socketError = WSAGetLastError();
            }
            *error = SocketError(L"connect", socketError);
            closesocket(socketHandle);
            return false;
        }
    }

    nonBlocking = 0;
    ioctlsocket(socketHandle, FIONBIO, &nonBlocking);
    const int sendTimeout = kSendTimeoutMs;
    const int receiveTimeout = kReceiveTimeoutMs;
    setsockopt(socketHandle, SOL_SOCKET, SO_SNDTIMEO,
               reinterpret_cast<const char*>(&sendTimeout),
               sizeof(sendTimeout));
    setsockopt(socketHandle, SOL_SOCKET, SO_RCVTIMEO,
               reinterpret_cast<const char*>(&receiveTimeout),
               sizeof(receiveTimeout));
    *connectedSocket = socketHandle;
    return true;
}

bool ExchangeWithServer(const std::wstring& query,
                        std::wstring* responseText,
                        std::wstring* error)
{
    WSADATA winsockData;
    const int startupStatus = WSAStartup(MAKEWORD(2, 2), &winsockData);
    if (startupStatus != 0) {
        *error = SocketError(L"WSAStartup", startupStatus);
        return false;
    }

    SOCKET socketHandle = INVALID_SOCKET;
    bool succeeded = ConnectToServer(
        kDefaultServerHost, kDefaultServerPort, &socketHandle, error);
    if (succeeded) {
        const std::string body =
            std::string("{\"text\":\"") +
            EscapeJsonString(WideToUtf8(query)) + "\"}";
        char header[512];
        _snprintf(header, sizeof(header),
                  "POST /message HTTP/1.1\r\n"
                  "Host: %s:%u\r\n"
                  "Content-Type: application/json; charset=utf-8\r\n"
                  "Content-Length: %u\r\n"
                  "Connection: close\r\n\r\n",
                  kDefaultServerHost,
                  static_cast<unsigned int>(kDefaultServerPort),
                  static_cast<unsigned int>(body.size()));
        header[sizeof(header) - 1] = '\0';
        succeeded = SendAll(socketHandle, std::string(header) + body, error);
    }

    std::string response;
    while (succeeded) {
        char buffer[4096];
        const int received = recv(socketHandle, buffer, sizeof(buffer), 0);
        if (received == 0) {
            break;
        }
        if (received == SOCKET_ERROR) {
            *error = SocketError(L"receive", WSAGetLastError());
            succeeded = false;
            break;
        }
        response.append(buffer, static_cast<size_t>(received));
        if (response.size() > kMaximumResponseBytes) {
            *error = L"server response exceeded 64 KiB";
            succeeded = false;
            break;
        }
    }

    if (socketHandle != INVALID_SOCKET) {
        closesocket(socketHandle);
    }
    WSACleanup();

    if (!succeeded) {
        return false;
    }
    const size_t bodyOffset = response.find("\r\n\r\n");
    if (bodyOffset == std::string::npos) {
        *error = L"server returned an invalid HTTP response";
        return false;
    }
    int statusCode = 0;
    if (sscanf(response.c_str(), "HTTP/%*u.%*u %d", &statusCode) != 1) {
        *error = L"server returned an invalid HTTP status";
        return false;
    }
    const std::string responseBody = response.substr(bodyOffset + 4);
    if (statusCode != 200) {
        if (ExtractJsonStringProperty(responseBody, "error", error)) {
            return false;
        }
        wchar_t statusMessage[96];
        _snwprintf(statusMessage,
                   sizeof(statusMessage) / sizeof(statusMessage[0]),
                   L"server returned HTTP status %d", statusCode);
        statusMessage[(sizeof(statusMessage) /
                       sizeof(statusMessage[0])) - 1] = L'\0';
        *error = statusMessage;
        return false;
    }
    if (!ExtractJsonStringProperty(responseBody, "text", responseText)) {
        *error = L"server response did not contain valid text";
        return false;
    }
    return true;
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
          oldEditorProc_(NULL), networkPending_(false),
          responseReady_(false), responseSucceeded_(false),
          thinkingAnimationActive_(false)
    {
        InitializeCriticalSection(&networkLock_);
        InterlockedIncrement(&g_objectCount);
    }

    virtual ~ClippyAddIn()
    {
        Stop();
        DeleteCriticalSection(&networkLock_);
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
        StopThinkingAnimation();
        if (g_activeAddIn == this) {
            g_activeAddIn = NULL;
        }
        if (application_ != NULL) {
            application_->Release();
            application_ = NULL;
        }
        EnterCriticalSection(&networkLock_);
        responseReady_ = false;
        responseText_.clear();
        responseError_.clear();
        LeaveCriticalSection(&networkLock_);
    }

    void Poll()
    {
        bool responseReady = false;
        bool responseSucceeded = false;
        std::wstring responseText;
        std::wstring responseError;
        EnterCriticalSection(&networkLock_);
        if (responseReady_) {
            responseReady = true;
            responseSucceeded = responseSucceeded_;
            responseText = responseText_;
            responseError = responseError_;
            responseReady_ = false;
            networkPending_ = false;
            responseText_.clear();
            responseError_.clear();
        }
        LeaveCriticalSection(&networkLock_);
        if (responseReady) {
            StopThinkingAnimation();
            if (responseSucceeded) {
                AppendLog(L"RESPONSE text=" + responseText);
                ShowResponse(L"Clippy host replied:", responseText);
            } else {
                AppendLog(L"ERROR host request failed: " + responseError);
                ShowResponse(L"Clippy couldn't answer:",
                             responseError);
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
        EnterCriticalSection(&networkLock_);
        const bool alreadyPending = networkPending_;
        LeaveCriticalSection(&networkLock_);
        if (alreadyPending || editor_ == NULL) {
            return alreadyPending;
        }
        const std::wstring query = Trim(GetWindowTextString(editor_));
        if (query.empty() || query == kPlaceholder) {
            return false;
        }

        AppendLog(std::wstring(L"QUERY source=") + source + L" text=" + query);
        EnterCriticalSection(&networkLock_);
        networkPending_ = true;
        responseReady_ = false;
        LeaveCriticalSection(&networkLock_);

        SetWindowTextW(editor_, L"Asking the Clippy host...");
        StartThinkingAnimation();

        NetworkJob* job = new NetworkJob();
        job->owner = this;
        job->query = query;
        AddRef();
        HANDLE thread = CreateThread(
            NULL, 0, NetworkThreadProc, job, 0, NULL);
        if (thread == NULL) {
            const DWORD error = GetLastError();
            delete job;
            Release();
            EnterCriticalSection(&networkLock_);
            networkPending_ = false;
            LeaveCriticalSection(&networkLock_);
            StopThinkingAnimation();
            wchar_t message[128];
            _snwprintf(message, sizeof(message) / sizeof(message[0]),
                       L"ERROR CreateThread failed error=%lu",
                       static_cast<unsigned long>(error));
            message[(sizeof(message) / sizeof(message[0])) - 1] = L'\0';
            AppendLog(message);
            SetWindowTextW(editor_, L"Could not start the host request.");
            return true;
        }
        CloseHandle(thread);
        return true;
    }

    void StartThinkingAnimation()
    {
        IDispatch* assistant = NULL;
        HRESULT hr = GetDispatchProperty(application_, L"Assistant", &assistant);
        if (SUCCEEDED(hr)) {
            hr = PutBool(assistant, L"Visible", true);
        }
        if (SUCCEEDED(hr)) {
            hr = PutLong(assistant, L"Animation", kMsoAnimationThinking);
        }
        if (assistant != NULL) {
            assistant->Release();
        }

        thinkingAnimationActive_ = SUCCEEDED(hr);
        if (thinkingAnimationActive_) {
            AppendLog(L"ANIMATION started Thinking");
        } else {
            wchar_t error[80];
            _snwprintf(error, sizeof(error) / sizeof(error[0]),
                       L"ERROR thinking animation failed hr=0x%08lX",
                       static_cast<unsigned long>(hr));
            error[(sizeof(error) / sizeof(error[0])) - 1] = L'\0';
            AppendLog(error);
        }
    }

    void StopThinkingAnimation()
    {
        if (!thinkingAnimationActive_) {
            return;
        }

        IDispatch* assistant = NULL;
        HRESULT hr = GetDispatchProperty(application_, L"Assistant", &assistant);
        if (SUCCEEDED(hr)) {
            hr = PutLong(assistant, L"Animation", kMsoAnimationIdle);
        }
        if (assistant != NULL) {
            assistant->Release();
        }
        thinkingAnimationActive_ = false;

        if (SUCCEEDED(hr)) {
            AppendLog(L"ANIMATION stopped Thinking");
        } else {
            wchar_t error[80];
            _snwprintf(error, sizeof(error) / sizeof(error[0]),
                       L"ERROR stopping thinking animation failed hr=0x%08lX",
                       static_cast<unsigned long>(hr));
            error[(sizeof(error) / sizeof(error[0])) - 1] = L'\0';
            AppendLog(error);
        }
    }

    static DWORD WINAPI NetworkThreadProc(LPVOID parameter)
    {
        NetworkJob* job = reinterpret_cast<NetworkJob*>(parameter);
        std::wstring responseText;
        std::wstring error;
        const bool succeeded = ExchangeWithServer(
            job->query, &responseText, &error);
        job->owner->CompleteNetworkRequest(
            succeeded, responseText, error);
        ClippyAddIn* owner = job->owner;
        delete job;
        owner->Release();
        return 0;
    }

    void CompleteNetworkRequest(bool succeeded,
                                const std::wstring& responseText,
                                const std::wstring& error)
    {
        EnterCriticalSection(&networkLock_);
        responseSucceeded_ = succeeded;
        responseText_ = responseText;
        responseError_ = error;
        responseReady_ = true;
        LeaveCriticalSection(&networkLock_);
    }

    void ShowResponse(const std::wstring& heading,
                      const std::wstring& text)
    {
        const BalloonMarkup markup = FormatBalloonMarkup(text);
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
            hr = PutString(balloon, L"Heading", heading);
            if (SUCCEEDED(hr)) {
                hr = PutString(balloon, L"Text", markup.text);
            }
            if (SUCCEEDED(hr) && !markup.labels.empty()) {
                hr = PutLong(balloon, L"BalloonType", markup.listType);
            }
            IDispatch* labels = NULL;
            if (SUCCEEDED(hr) && !markup.labels.empty()) {
                hr = GetDispatchProperty(balloon, L"Labels", &labels);
            }
            for (size_t index = 0;
                 SUCCEEDED(hr) && index < markup.labels.size(); ++index) {
                IDispatch* label = NULL;
                hr = GetIndexedDispatch(
                    labels, static_cast<LONG>(index + 1), &label);
                if (SUCCEEDED(hr)) {
                    hr = PutString(label, L"Text", markup.labels[index]);
                }
                if (label != NULL) {
                    label->Release();
                }
            }
            if (labels != NULL) {
                labels->Release();
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
            AppendLog(L"RESPONSE shown with Assistant.NewBalloon");
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
    struct NetworkJob {
        ClippyAddIn* owner;
        std::wstring query;
    };
    CRITICAL_SECTION networkLock_;
    bool networkPending_;
    bool responseReady_;
    bool responseSucceeded_;
    bool thinkingAnimationActive_;
    std::wstring responseText_;
    std::wstring responseError_;
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
