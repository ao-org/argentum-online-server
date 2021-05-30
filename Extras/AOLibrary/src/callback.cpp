#include <declares.h>
#include <stdint.h>

EXPORT void CALLBACK CallHandle(DWORD address, int16_t userindex) {
    if (userindex) {
        typedef void(CALLBACK* SubUI)(int16_t userindex);
        SubUI FunctionCall = reinterpret_cast<SubUI>(address);
        FunctionCall(userindex);
    } else {
        typedef void(CALLBACK* SubSinUI)();
        SubSinUI FunctionCall = reinterpret_cast<SubSinUI>(address);
        FunctionCall();
    }
}
