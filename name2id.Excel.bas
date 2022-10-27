Attribute VB_Name = "HashFNV1a"
Option Explicit

'*******************************************************************************
' Module:   HashFNV1a
' Version:  1.0
' Created:  20-10-2022
' Authors:  Leonardo Gama
' Homepage: github.com/leogama/spreadsheet-hash
'
' Description
' -----------
'
' This BASIC module implements the Fowler-Noll-Vo (FNV) hash function without
' dependencies on external libraries.  Specifically, it implements the FNV1a
' 32 bits variant, outputting integers of sizes between 32 and 4 bits by
' applying the XOR-folding technique.
'
' Public functions:
'   HASH(text as String, Optional length as Long, Optional salt as String) as String
'   NAME2ID(text as String, Optional length as Long, Optional salt as String) as String
'
' MIT License
' -----------
'
' Copyright  2022  Leonardo dos Reis Gama
'
' Permission is hereby granted, free of charge, to any person obtaining a
' copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
'*******************************************************************************

Sub Main()
    ' Check that a salt text is just prepended to the input text and
    ' that full name normalization works.
    Dim hash1 as String, hash2 as String, id1 as String, id2 as String
    hash1 = HASH("hello world")
    hash2 = HASH("world", salt:="hello ")
    id1 = String(16, "*")
    RSet id1 = NAME2ID("Maria  D.AssunÁ„o ", 4)
    id2 = NAME2ID("MARIA D ASSUNCAO", 4)
    MsgBox hash1 & " == " & hash2 & CHR$(10) & id1 & " == " & id2
End Sub

Public Function HASH(text as String, Optional length as Variant, Optional salt as Variant)
    ' Calculate a hash for the text input based on the FNV hash function.
    '
    ' Generate an hexadecimal hash with length between 8 and 1 from the
    ' 32 bit version of the FNV1a hash.  The hash is shortened to the
    ' specified length by XOR-folding.  The salt text, if specified, is
    ' prepend to the input text.
    '
    ' Parameters:
    '   text [String]: the input text
    '   length [Long]: length of the hexadecimal output (optional; default: 8)
    '   salt [String]: secret salt text (optional)

    Const MAX_LENGTH as Long = 8  ' 32 bits in half-bytes 
    Const MIN_LENGTH as Long = 1  ' 4 bits (pretty useless, but allowed)
    Dim hashHi as String, hashLo as String

    If Not IsMissing(salt) Then
        text = salt & text
    End If
    
    If IsMissing(length) Then
        length = MAX_LENGTH
    ElseIf IsNumeric(length) Then
        length = CLng(length)
        If length = 0 Then length = MAX_LENGTH  ' LibreOffice missing argument bug
        If length < MIN_LENGTH Or length > MAX_LENGTH Then GoTo ValueError
    Else
        GoTo ValueError
    End If

    HASH = FNV1a(text)
    
    ' Reduce hash length by XOR-folding.
    If length < MAX_LENGTH Then
        If length >= 4 Then
            ' XOR the highest bits with the lowest bits.
            hashHi = Left(HASH, length)
            hashLo = Right(HASH, MAX_LENGTH - length)
        Else
            ' XOR the lowest n bits with the second-lowest n bits.
            hashHi = Mid(HASH, 1 + MAX_LENGTH - 2 * length, length)  ' drop higher bits
            hashLo = Right(HASH, length)
        End If
        HASH = Hex(CLng("&H" & hashHi) Xor CLng("&H" & hashLo))
        HASH = String(length - Len(HASH), "0") & HASH
    End If
    Exit Function

ValueError:
    On Error GoTo NotExcel
    HASH = CVErr(xlErrValue)  ' #VALUE!
    Exit Function
NotExcel:
    HASH = CVErr(502)  ' invalid argument
    Exit Function

End Function

Public Function NAME2ID(fullName as String, Optional length as Variant, Optional salt as Variant)
    ' Calculate a hash for a person's name after normalizing it.
    '
    ' Generate an hexadecimal hash from a full name after normalizing it
    ' to guarantee consistency for non-standardized input.  The full
    ' name is processed, before being passed to the function HASH,
    ' by following these steps:
    '   1. It is converted to all uppercase letters;
    '   2. Leading and trailing whitespaces are stripped;
    '   3. Periods from abbreviations are removed;
    '   4. Multiple whitespaces between names are replaced by a single one;
    '   5. Accented latin letters (from the Latin-1 encoding) are replaced
    '   by their unaccented versions.
    '
    ' As an example, the input "Maria  D. ConceiÁ„o " would be normalized
    ' to "MARIA D CONCEICAO".
    '
    ' Parameters:
    '   text [String]: the input text
    '   length [Long]: length of the hexadecimal output (optional; default: 8)
    '   salt [String]: secret salt text (optional)

    Const ACCENTED as String = "¿¡¬√ƒ≈«»… ÀÃÕŒœ—“”‘’÷Ÿ⁄€‹›"  ' from Latin-1 encoding
    Const UNACCENTED as String = "AAAAAACEEEEIIIINOOOOOUUUUY"
    Dim text as String, i as Long

    ' Trim whitespaces, set to uppercase and remove periods from abbreviations.
    text = Replace(Trim(UCase(fullName)), ".", " ")
    ' Remove double whitespaces from the middle.
    Do While InStr(text, "  ")
        text = Replace(text, "  ", " ")
    Loop
    ' Replace accented latin letters by unnacented ones.
    For i = 1 To Len(ACCENTED)
        text = Replace(text, Mid(ACCENTED, i, 1), Mid(UNACCENTED, i, 1))
    Next i

    NAME2ID = HASH(text, length, salt)
End Function

Private Function FNV1a(text as String) as String
    ' Implements the FNV1a 32 bit hash function variant.
    '
    ' Characters are fed to the hash loop with the sequence of bytes from their
    ' UTF-8 representation.  The output is in length 8 hexadecimal form.

    ' FNV offset basis
    Const FNV32_BASIS as Long = &H811C9DC5
    Const HASH_LENGTH as Long = 8  ' 32 bits in half-bytes
    ' UTF-8 conversion
    Const UTF8_1_MAX as Long = &H007F
    Const UTF8_2_MAX as Long = &H07FF
    Const UTF8_2_BYTES_PATTERN as Long = &HC0
    Const UTF8_3_BYTES_PATTERN as Long = &HE0
    Const UTF8_CONTIN_PATTERN as Long = &H80
    Const RSHIFT_6_BITS as Long = 2 ^ 6
    Const RSHIFT_12_BITS as Long = 2 ^ 12
    Const MASK_6_BITS as Long = &H3F
    Dim hashLng as Long, hashDbl as Double, codepoint as Long, i as Long

    hashLng = FNV32_BASIS
    For i = 1 To Len(text)
        codepoint = Asc(Mid(text, i, 1))

        ' Split into 1, 2 or 3-bytes UTF-8 form depending on the codepoint value.
        If codepoint <= UTF8_1_MAX Then
            hashLng = IterateFNV1a(hashLng, codepoint)
        ElseIF codepoint <= UTF8_2_MAX Then
            hashLng = IterateFNV1a(hashLng, (codepoint \ RSHIFT_6_BITS) Or UTF8_2_BYTES_PATTERN)
            hashLng = IterateFNV1a(hashLng, (codepoint And MASK_6_BITS) Or UTF8_CONTIN_PATTERN)
        Else
            hashLng = IterateFNV1a(hashLng, (codepoint \ RSHIFT_12_BITS) Or UTF8_3_BYTES_PATTERN)
            hashLng = IterateFNV1a(hashLng, ((codepoint \ RSHIFT_6_BITS) And MASK_6_BITS) Or UTF8_CONTIN_PATTERN)
            hashLng = IterateFNV1a(hashLng, (codepoint And MASK_6_BITS) Or UTF8_CONTIN_PATTERN)
        End If
    Next i

    FNV1a = Hex(hashLng)
    FNV1a = String(HASH_LENGTH - Len(FNV1a), "0") & FNV1a
End Function

Private Function IterateFNV1a(hashLng as Long, textByte as Long) as Long
    ' Loop of the FNV1a algorithm.
    ' 
    ' Only integer types can do bitwise operations and only Double can do the
    ' multiplication step without overflow.  The intermediary results must be
    ' converted between each step.  The sign bit of the Long type must be
    ' handled specially as the algorithm requires unsigned integer arithmetic.
    '
    ' Even if the Double type can store up to 53 bits, the product of the hash
    ' by the FNV prime can reach 61 bits.  Therefore, it's necessery to do a
    ' piecewise multiplication with the higher and lower parts of the prime.
    ' The lower product has a maximum size (for 0h193 * 0hFFFFFFFF) of 41 bits.
    ' The higher product is calculated by a "left shift" of the 8 lowest bits
    ' of the hash, moving them to most significant position of the 32 bit word.
    ' The partial products are then summed and truncated to 32 bits.

    ' FNV prime
    Const FNV32_PRIME_LO as Double = 403     ' lower bits of 0x01000193
    Const FNV32_PRIME_HI as Double = 2 ^ 24  ' higher bits of 0x01000193
    Const PRIME_MASK as Long = 2 ^ 8 - 1     ' 8 bits = 32 bits - 24 bits
    ' Sign bit handling in type conversion
    Const MASK_SIGN as Long = &H7FFFFFFF
    Const SIGN_BIT_LONG as Long = &H80000000
    Const SIGN_BIT_DOUBLE as Double = 2 ^ 31
    Dim hashDbl as Double, hashDblHi as Double

    ' Step 1: XOR with input
    hashLng = hashLng Xor textByte
    
    ' Convert to Double
    hashDbl = CDbl(hashLng And MASK_SIGN)  ' strip sign bit
    If hashLng < 0 Then hashDbl = hashDbl + SIGN_BIT_DOUBLE

    ' Step 2: multiply with prime
    hashDbl = hashDbl * FNV32_PRIME_LO + (hashLng And PRIME_MASK) * FNV32_PRIME_HI

    ' Convert to Long (bits above the 32nd are ignored)
    hashDblHi = Fix(hashDbl / SIGN_BIT_DOUBLE)  ' poor man's module
    hashLng = CLng(hashDbl - SIGN_BIT_DOUBLE * hashDblHi)
    If hashDblHi And 1 Then hashLng = hashLng Or SIGN_BIT_LONG

    IterateFNV1a = hashLng
End Function
