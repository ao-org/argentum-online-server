#pragma once

#include <Windows.h>

#define WIN32_LEAN_AND_MEAN // Excluir material rara vez utilizado de encabezados de Windows
#define VC_EXTRALEAN        // Excluir material rara vez utilizado de encabezados de Windows (de MSVC)

#define EXPORT extern "C" __declspec(dllexport)