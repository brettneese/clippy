#define _WIN32_WINNT 0x0501
#define WINVER 0x0501

#include <windows.h>
#include <oleauto.h>
#include <stdio.h>

namespace {

void PrintFailure(const wchar_t* operation, HRESULT hr)
{
    wchar_t* systemMessage = NULL;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                       FORMAT_MESSAGE_FROM_SYSTEM |
                       FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL,
                   hr,
                   0,
                   reinterpret_cast<wchar_t*>(&systemMessage),
                   0,
                   NULL);

    if (systemMessage != NULL) {
        fwprintf(stderr, L"ERROR: %s failed (0x%08lX): %s\n",
                 operation, static_cast<unsigned long>(hr), systemMessage);
        LocalFree(systemMessage);
    } else {
        fwprintf(stderr, L"ERROR: %s failed (0x%08lX).\n",
                 operation, static_cast<unsigned long>(hr));
    }
}

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
        PrintFailure(memberName, hr);
        return hr;
    }

    DISPID propertyPutId = DISPID_PROPERTYPUT;
    DISPPARAMS parameters;
    ZeroMemory(&parameters, sizeof(parameters));
    parameters.rgvarg = arguments;
    parameters.cArgs = argumentCount;
    if ((flags & DISPATCH_PROPERTYPUT) != 0) {
        parameters.rgdispidNamedArgs = &propertyPutId;
        parameters.cNamedArgs = 1;
    }

    VARIANT localResult;
    VariantInit(&localResult);
    if (result == NULL) {
        result = &localResult;
    } else {
        VariantInit(result);
    }

    EXCEPINFO exceptionInfo;
    ZeroMemory(&exceptionInfo, sizeof(exceptionInfo));
    UINT argumentError = 0;
    hr = object->Invoke(memberId, IID_NULL, LOCALE_USER_DEFAULT, flags,
                        &parameters, result, &exceptionInfo, &argumentError);
    if (FAILED(hr)) {
        fwprintf(stderr, L"ERROR: invoking %s failed (0x%08lX)",
                 memberName, static_cast<unsigned long>(hr));
        if (exceptionInfo.bstrDescription != NULL) {
            fwprintf(stderr, L": %s", exceptionInfo.bstrDescription);
        }
        fwprintf(stderr, L".\n");
    }

    SysFreeString(exceptionInfo.bstrSource);
    SysFreeString(exceptionInfo.bstrDescription);
    SysFreeString(exceptionInfo.bstrHelpFile);
    if (result == &localResult) {
        VariantClear(&localResult);
    }
    return hr;
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

HRESULT GetDispatchProperty(IDispatch* object,
                            const wchar_t* memberName,
                            IDispatch** value)
{
    if (value == NULL) {
        return E_POINTER;
    }
    *value = NULL;

    VARIANT result;
    VariantInit(&result);
    HRESULT hr = InvokeByName(object, memberName, DISPATCH_PROPERTYGET,
                              NULL, 0, &result);
    if (SUCCEEDED(hr)) {
        if (result.vt == VT_DISPATCH && result.pdispVal != NULL) {
            *value = result.pdispVal;
            (*value)->AddRef();
        } else if (result.vt == VT_UNKNOWN && result.punkVal != NULL) {
            hr = result.punkVal->QueryInterface(IID_IDispatch,
                                                reinterpret_cast<void**>(value));
        } else {
            hr = DISP_E_TYPEMISMATCH;
        }
    }
    VariantClear(&result);
    return hr;
}

HRESULT CallWithStrings(IDispatch* object,
                        const wchar_t* memberName,
                        const wchar_t** values,
                        UINT valueCount)
{
    VARIANT arguments[2];
    if (valueCount > 2) {
        return E_INVALIDARG;
    }

    UINT index;
    for (index = 0; index < valueCount; ++index) {
        VariantInit(&arguments[index]);
        arguments[index].vt = VT_BSTR;
        arguments[index].bstrVal = SysAllocString(values[valueCount - index - 1]);
        if (arguments[index].bstrVal == NULL) {
            while (index > 0) {
                --index;
                VariantClear(&arguments[index]);
            }
            return E_OUTOFMEMORY;
        }
    }

    HRESULT hr = InvokeByName(object, memberName, DISPATCH_METHOD,
                              arguments, valueCount, NULL);
    for (index = 0; index < valueCount; ++index) {
        VariantClear(&arguments[index]);
    }
    return hr;
}

HRESULT GetNamedDispatch(IDispatch* object,
                         const wchar_t* memberName,
                         const wchar_t* name,
                         IDispatch** value)
{
    if (value == NULL) {
        return E_POINTER;
    }
    *value = NULL;

    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_BSTR;
    argument.bstrVal = SysAllocString(name);
    if (argument.bstrVal == NULL) {
        return E_OUTOFMEMORY;
    }

    VARIANT result;
    VariantInit(&result);
    HRESULT hr = InvokeByName(object, memberName,
                              DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                              &argument, 1, &result);
    VariantClear(&argument);
    if (SUCCEEDED(hr)) {
        if (result.vt == VT_DISPATCH && result.pdispVal != NULL) {
            *value = result.pdispVal;
            (*value)->AddRef();
        } else if (result.vt == VT_UNKNOWN && result.punkVal != NULL) {
            hr = result.punkVal->QueryInterface(IID_IDispatch,
                                                reinterpret_cast<void**>(value));
        } else {
            hr = DISP_E_TYPEMISMATCH;
        }
    }
    VariantClear(&result);
    return hr;
}

}  // namespace

int wmain()
{
    const wchar_t* characterName = L"Clippy";
    const wchar_t* characterPath =
        L"C:\\Program Files\\Microsoft Office\\Office10\\CLIPPIT.ACS";

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        PrintFailure(L"CoInitializeEx", hr);
        return 10;
    }

    int exitCode = 0;
    IDispatch* agent = NULL;
    IDispatch* characters = NULL;
    IDispatch* clippy = NULL;

    CLSID controlClass;
    hr = CLSIDFromProgID(L"Agent.Control.2", &controlClass);
    if (FAILED(hr)) {
        PrintFailure(L"CLSIDFromProgID(Agent.Control.2)", hr);
        exitCode = 11;
        goto Cleanup;
    }

    hr = CoCreateInstance(controlClass, NULL, CLSCTX_INPROC_SERVER,
                          IID_IDispatch, reinterpret_cast<void**>(&agent));
    if (FAILED(hr)) {
        PrintFailure(L"CoCreateInstance(Agent.Control.2)", hr);
        exitCode = 12;
        goto Cleanup;
    }
    wprintf(L"OK: created Agent.Control.2.\n");

    hr = PutBool(agent, L"Connected", true);
    if (FAILED(hr)) {
        exitCode = 13;
        goto Cleanup;
    }
    wprintf(L"OK: connected to Microsoft Agent.\n");

    hr = GetDispatchProperty(agent, L"Characters", &characters);
    if (FAILED(hr)) {
        PrintFailure(L"Agent.Characters", hr);
        exitCode = 14;
        goto Cleanup;
    }

    {
        const wchar_t* loadArguments[] = {characterName, characterPath};
        hr = CallWithStrings(characters, L"Load", loadArguments, 2);
    }
    if (FAILED(hr)) {
        exitCode = 15;
        goto Cleanup;
    }
    wprintf(L"OK: loaded %s.\n", characterPath);

    Sleep(500);
    hr = GetNamedDispatch(characters, L"Character", characterName, &clippy);
    if (FAILED(hr)) {
        PrintFailure(L"Characters.Character(Clippy)", hr);
        exitCode = 16;
        goto Cleanup;
    }

    hr = InvokeByName(clippy, L"Show", DISPATCH_METHOD, NULL, 0, NULL);
    if (FAILED(hr)) {
        exitCode = 17;
        goto Cleanup;
    }
    wprintf(L"OK: invoked Clippy.Show().\n");

    {
        const wchar_t* thinkArguments[] = {
            L"Native C++ reached Microsoft Agent successfully."
        };
        hr = CallWithStrings(clippy, L"Think", thinkArguments, 1);
    }
    if (FAILED(hr)) {
        exitCode = 18;
        goto Cleanup;
    }
    wprintf(L"OK: invoked Clippy.Think().\n");

    Sleep(3000);

Cleanup:
    if (clippy != NULL) {
        clippy->Release();
    }
    if (characters != NULL) {
        characters->Release();
    }
    if (agent != NULL) {
        agent->Release();
    }
    CoUninitialize();
    return exitCode;
}
