// Utf8Conv.h
//
// Source code file modified from:
// https://github.com/GiovanniDicanio/Utf8ConvAtlStl
// Copyright (c) 2016 by Giovanni Dicanio
// SPDX-License-Identifier: MIT
//
// Unicode UTF-8 <-> UTF-16 String Conversion Functions

#pragma once

#include <Windows.h>    // Win32 Platform SDK main header        

#include <stdexcept>    // For std::runtime_error
#include <string>       // For std::string (UTF-8)

//==============================================================================
//                          Implementations
//==============================================================================


//------------------------------------------------------------------------------
// Exception class representing a conversion error between UTF-8 and UTF-16.
//------------------------------------------------------------------------------
class Utf8ConversionException
    : public std::runtime_error
{
public:

    // Create the exception object from error message and error code.
    Utf8ConversionException(const char* message, DWORD errorCode)
        : std::runtime_error(message)
        , m_errorCode(errorCode)
    {}

    // Create the exception object from error message and error code.
    Utf8ConversionException(const std::string& message, DWORD errorCode)
        : std::runtime_error(message)
        , m_errorCode(errorCode)
    {}

    // Conversion error code (as returned by GetLastError).
    DWORD ErrorCode() const
    {
        return m_errorCode;
    }


    // *** PRIVATE IMPLEMENTATION ***

private:
    // Error code as returned by GetLastError
    DWORD m_errorCode;
};


//------------------------------------------------------------------------------
// Convert form UTF-16 to UTF-8.
//
// UTF-16 strings are specified passing a BSTR.
// UTF-8 strings are stored using std::string.
// 
// On conversion errors (e.g. invalid UTF-16 sequence in input string), throws
// Utf8ConversionException.
//------------------------------------------------------------------------------
class Utf8FromUtf16
    : public std::string //lint -e1509 base class destructor for class '...' is not virtual
{
public:
    explicit Utf8FromUtf16(BSTR utf16Start)
    {
        size_t const utf16LengthUsingSizet = SysStringLen(utf16Start);

        // Special case of empty input
        if (utf16LengthUsingSizet == 0)
        {
            // Empty input ==> empty result
            return;
        }

        // Safely cast the length of the source UTF-16 string (expressed in wchar_ts)
        // from size_t to int for the WideCharToMultiByte API.
        // If the size_t value is too big to be stored into an int, 
        // throw an exception to prevent conversion errors (bugs) like huge size_t values 
        // converted to *negative* integers.
        if (utf16LengthUsingSizet > INT_MAX)
        {
            throw Utf8ConversionException(
                "Input string too long: size_t-length doesn't fit into int.\n",
                ERROR_INVALID_PARAMETER);
        }

        // Length of source string view, in wchar_ts
        int const utf16Length = static_cast<int>(utf16LengthUsingSizet);

        // Get the length, in chars, of the resulting UTF-8 string
        int const utf8Length = ::WideCharToMultiByte(
            CP_UTF8,            // convert to UTF-8
            0,                  // conversion flags
            utf16Start,         // source UTF-16 string
            utf16Length,        // length of source UTF-16 string, in wchar_ts
            NULL,               // unused - no conversion required in this step
            0,                  // request size of destination buffer, in chars
            NULL, NULL          // unused
        );
        if (utf8Length == 0)
        {
            // Conversion error: capture error code and throw
            DWORD const error = ::GetLastError();
            throw Utf8ConversionException(
                "Error in converting from UTF-16 to UTF-8.\n",
                error);
        }

        // Make room in the destination string for the converted bits
        resize(utf8Length);

        // Do the actual conversion from UTF-16 to UTF-8
        int const result = ::WideCharToMultiByte(
            CP_UTF8,            // convert to UTF-8
            0,                  // conversion flags
            utf16Start,         // source UTF-16 string
            utf16Length,        // length of source UTF-16 string, in wchar_ts
            &*begin(),          // pointer to destination buffer
            utf8Length,         // size of destination buffer, in chars
            NULL, NULL          // unused
        );
        if (result == 0)
        {
            // Conversion error: capture error code and throw
            DWORD const error = ::GetLastError();
            throw Utf8ConversionException(
                "Error in converting from UTF-16 to UTF-8.\n",
                error);
        }
    }
};

//------------------------------------------------------------------------------
// Convert form UTF-8 to UTF-16.
//
// UTF-8 strings are specified using start and length.
// UTF-16 strings are stored in BSTR.
// 
// On conversion errors (e.g. invalid UTF-8 sequence in input string), returns
// HRESULT.
//------------------------------------------------------------------------------
inline HRESULT Utf16FromUtf8(const char *utf8Start, size_t const utf8LengthUsingSizet, BSTR *utf16)
{
    // Special case of empty input
    if (utf8LengthUsingSizet == 0)
    {
        // Empty input ==> empty result
        SysReAllocStringLen(utf16, NULL, 0);
        return S_OK;
    }

    // Safely cast the length of the source UTF-8 string (expressed in chars)
    // from size_t to int for the MultiByteToWideChar API.
    // If the size_t value is too big to be stored into an int, 
    // throw an exception to prevent conversion errors (bugs) like huge size_t values 
    // converted to *negative* integers.
    if (utf8LengthUsingSizet > INT_MAX)
    {
        return E_OUTOFMEMORY;
    }

    int const utf8Length = static_cast<int>(utf8LengthUsingSizet);

    // Get the size of the destination UTF-16 string
    int const utf16Length = ::MultiByteToWideChar(
        CP_UTF8,       // source string is in UTF-8
        0,             // conversion flags
        utf8Start,     // source UTF-8 string pointer
        utf8Length,    // length of the source UTF-8 string, in chars
        NULL,          // unused - no conversion done in this step
        0              // request size of destination buffer, in wchar_ts
    );
    if (utf16Length == 0)
    {
        // Conversion error: capture error code and throw
        DWORD const error = ::GetLastError();
        return HRESULT_FROM_WIN32(error);
    }

    // Make room in the destination string for the converted bits
    if (!SysReAllocStringLen(utf16, NULL, utf16Length))
    {
        return E_OUTOFMEMORY;
    }

    wchar_t *const utf16Buffer = *utf16;

    // Do the actual conversion from UTF-8 to UTF-16
    int result = ::MultiByteToWideChar(
        CP_UTF8,       // source string is in UTF-8
        0,             // conversion flags
        utf8Start,     // source UTF-8 string pointer
        utf8Length,    // length of source UTF-8 string, in chars
        utf16Buffer,   // pointer to destination buffer
        utf16Length    // size of destination buffer, in wchar_ts           
    );
    if (result == 0)
    {
        // Conversion error: capture error code and throw
        DWORD const error = ::GetLastError();
        return HRESULT_FROM_WIN32(error);
    }
    return S_OK;
}


//------------------------------------------------------------------------------
// Convert form UTF-8 to UTF-16.
//
// UTF-8 strings are stored using std::string.
// UTF-16 strings are stored in BSTR.
// 
// On conversion errors (e.g. invalid UTF-8 sequence in input string), returns
// HRESULT.
//------------------------------------------------------------------------------
inline HRESULT Utf16FromUtf8(const std::string& utf8, BSTR *utf16)
{
    return Utf16FromUtf8(utf8.data(), utf8.length(), utf16);
}
