#include <Windows.h>
#include <stdint.h>

extern "C" {
    void __declspec(dllexport) CALLBACK CallHandle(DWORD address, int16_t userindex) {
        typedef void(__stdcall* SubUI)(int16_t userindex);
        typedef void(__stdcall* SubSinUI)();

        if (userindex) {
            SubUI FunctionCall = reinterpret_cast<SubUI>(address);
            FunctionCall(userindex);
        } else {
            SubSinUI FunctionCall = reinterpret_cast<SubSinUI>(address);
            FunctionCall();
        }
    }
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    return TRUE;
}