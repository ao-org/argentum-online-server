# AO20CryptoSysWrapper Module Documentation

## 1. General Purpose

The `AO20CryptoSysWrapper.bas` module acts as a wrapper for cryptographic functionalities, primarily focusing on encryption and decryption services. It abstracts the direct calls to an underlying external cryptography library, likely the "CryptoSys API" (e.g., `diCryptoSys.dll`), to provide simplified encryption and decryption routines tailored for the game server's needs.

The core cryptographic operations performed are AES-128 encryption and decryption using the Cipher Feedback (CFB) mode with no padding. The module also includes various utility functions for data type conversions between strings, byte arrays, hexadecimal format, and Base64 encoding, which are necessary for preparing data for and processing data from the cryptographic functions.

## 2. Main Public Subroutines and Functions

*   **`Encrypt(hex_key As String, plain_text As String) As String`**:
    *   Encrypts the input `plain_text` string.
    *   Uses AES-128 in CFB mode with no padding.
    *   The encryption key is provided as a hexadecimal string (`hex_key`), which is also used as the Initialization Vector (IV).
    *   The input `plain_text` is first converted to its hexadecimal representation, then to bytes before encryption.
    *   The resulting encrypted byte array is then Base64 encoded to produce the final output string.

*   **`Decrypt(hex_key As String, encrypted_text_b64 As String) As String`**:
    *   Decrypts the input `encrypted_text_b64` string, which is expected to be Base64 encoded.
    *   Uses AES-128 in CFB mode with no padding.
    *   The decryption key is provided as a hexadecimal string (`hex_key`), also used as the IV.
    *   The Base64 input is first decoded, then converted to its hexadecimal representation, and finally to bytes before decryption.
    *   The decrypted byte array is converted back to a hexadecimal string and then to the original plain text string.

*   **`Str2ByteArr(str As String, arr() As Byte, Optional length As Long)`**:
    *   Converts a given string `str` into a byte array `arr`. If `length` is specified, the array is sized accordingly.

*   **`ByteArr2String(arr() As Byte) As String`**:
    *   Converts a byte array `arr` back into a string.

*   **`hiByte(w As Integer) As Byte`**:
    *   Extracts and returns the high byte from a 16-bit integer.

*   **`LoByte(w As Integer) As Byte`**:
    *   Extracts and returns the low byte from a 16-bit integer.

*   **`MakeInt(LoByte As Byte, hiByte As Byte) As Integer`**:
    *   Constructs a 16-bit integer from its constituent low and high bytes.

*   **`CopyBytes(src() As Byte, dst() As Byte, size As Long, Optional offset As Long)`**:
    *   Copies `size` bytes from the source byte array `src` to the destination byte array `dst`, starting at the given `offset` in the destination array if specified.

*   **`ByteArrayToHex(ByteArray() As Byte) As String`**:
    *   Converts a byte array into a string of hexadecimal values, with each byte representation separated by a space.

*   **`initBase64Chars()`**:
    *   Initializes a global array `base64_chars` containing the standard 64 characters used for Base64 encoding, plus the padding character '='. This is used by the custom `IsBase64` function.

*   **`IsBase64(str As String) As Boolean`**:
    *   Checks if the input string `str` contains only valid Base64 characters as defined in the `base64_chars` array.

## 3. Notable Dependencies

*   **External Cryptography Library (Implicit - likely CryptoSys API):**
    *   The module relies heavily on functions that are not standard to VB6 and are characteristic of the "CryptoSys API" by DI Management Services Pty Ltd. These functions include:
        *   `cnvBytesFromHexStr`
        *   `cnvHexStrFromString`
        *   `cipherEncryptBytes2`
        *   `cnvToBase64`
        *   `cnvFromBase64`
        *   `cnvToHex`
        *   `cnvStringFromHexStr`
        *   `cipherDecryptBytes2`
    *   These functions would be imported from an external Dynamic Link Library (DLL), such as `diCryptoSys.dll`. The module's name itself, `AO20CryptoSysWrapper`, strongly suggests this dependency.

*   **Global Variables:**
    *   `base64_chars(1 To 65) As String`: A module-level array used to store the character set for Base64 validation. Initialized by `initBase64Chars()`.

This module encapsulates the direct interaction with the external cryptography library, providing a simplified interface for encryption and decryption tasks within the server.
