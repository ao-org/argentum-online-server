#include <declares.h>
#include <stdint.h>

EXPORT void CALLBACK CallHandle(DWORD address, int16_t userindex) {
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