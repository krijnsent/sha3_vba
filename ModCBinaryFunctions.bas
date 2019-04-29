Attribute VB_Name = "ModCBinaryFunctions"
'INSPRIRED BY:
'https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Bitwise_Operators
'https://www.mrexcel.com/forum/excel-questions/578667-use-dec2bin-function-vba-edit-macro.html
'https://vbaf1.com/variables/data-types/

Sub TestBinaryFunctions()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModCBinaryFunctions"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestDec2Bin")
'8 bit
Test.IsEqual DecToBin_C(1, 8), "00000001"
Test.IsEqual DecToBin_C(-1, 8), "11111111"
Test.IsEqual DecToBin_C(127, 8), "01111111"
Test.IsEqual DecToBin_C(-128, 8), "10000000"
'16 bit
Test.IsEqual DecToBin_C(10, 16), "0000000000001010"
Test.IsEqual DecToBin_C(-10, 16), "1111111111110110"

'Tests for signed/unsigned
Test.IsEqual DecToBin_C(255, 8, False), "11111111"
Test.IsEqual DecToBin_C(255, 8, True), ""  'ERROR -> returns empty string


Set Test = Suite.Test("TestBin2Dec")
'8 bit
Test.IsEqual BinToDec_C("00000001"), 1
Test.IsEqual BinToDec_C("11111111"), -1
Test.IsEqual BinToDec_C("01111111"), 127
Test.IsEqual BinToDec_C("10000000"), -128
'16 bit
Test.IsEqual BinToDec_C("0000000000001010"), 10
Test.IsEqual BinToDec_C("1111111111110110"), -10

'Signed
Test.IsEqual BinToDec_C("10000000", False), 128
Test.IsEqual BinToDec_C("11111111", False), 255
'Signed Long
Test.IsEqual BinToDec_C("01111111111111111111111111111111", False), 2147483647
Test.IsEqual BinToDec_C("01111111111111111111111111111111", True), 2147483647
Test.IsEqual BinToDec_C("10000000000000000000000000000001", False), 2147483649#
Test.IsEqual BinToDec_C("10000000000000000000000000000001", True), -2147483647
'Unsigned 64 bit max
Test.IsEqual BinToDec_C("0111111111111111111111111111111111111111111111111111111111111111", False), 2 ^ (64 - 1) - 1


Set Test = Suite.Test("TestBinShift")
'Standard settings, default shift, small numbers will be signed bytes (-128 to +127)
Test.IsEqual LeftShift(1, 2), 4
Test.IsEqual LeftShift(-1, 2), -4

Test.IsEqual RightShift(1, 2), 0
Test.IsEqual RightShift(-1, 2), -1

Test.IsEqual RightShiftZF(1, 2), 0
Test.IsEqual RightShiftZF(-1, 2), 63

'integers
Test.IsEqual LeftShift(1000, 3), 8000
Test.IsEqual LeftShift(-1000, 3), -8000

'Integer shift falling over the edge
Test.IsEqual LeftShift(10000, 3), 14464
'10000 -> default 16-bit (integer): 0010011100010000 -> LeftShift(3) -> 0011100010000000 -> 14464
Test.IsEqual LeftShift(-10000, 3), -14464
'-10000 -> default 16-bit (integer): 1101100011110000 -> LeftShift(3) -> 1100011110000000 -> -14464
Test.IsEqual LeftShift(10000, 2), -25536
'Emptying with LeftShift (same length as nr of bits)
Test.IsEqual LeftShift(10000, 16), 0

'Test for options, e.g. Long (32-bits)
Test.IsEqual LeftShift(10000, 3, 32), 80000
'Test for unsigned value
Test.IsEqual LeftShift(64000, 3, 16, False), 53248
'64000 -> 16 bit unsigned -> 1111101000000000 -> LeftShift(3) -> 1101000000000000 -> Unsigned: 53248


Set Test = Suite.Test("TestHex2Dec")
'Test Hex values
Test.IsEqual HexToDec_C("0"), 0
Test.IsEqual HexToDec_C("f"), 15
Test.IsEqual HexToDec_C("ff"), 255
Test.IsEqual HexToDec_C("80000000"), 2147483648#
Test.IsEqual HexToDec_C("8000000a"), 2147483658#


Set Test = Suite.Test("TestXor")
Test.IsEqual 16 Xor 20, 4
Test.IsEqual Xor_C(16, 20), 4
Test.IsEqual -16 Xor 20, -28
Test.IsEqual Xor_C(-16, 20), -28
Test.IsEqual 15 Xor 15, 0
Test.IsEqual Xor_C(15, 15), 0

'Unsigned longs test
Test.IsEqual Xor_C(2606951925#, 32906, False), 2606919039#
Test.IsEqual Xor_C(2264371258#, 2147516416#, False), 116854842
Test.IsEqual Xor_C(1823284833, 2147483658#, False), 3970768491#

Set Test = Suite.Test("TestOr")
Test.IsEqual 3 Or 20, 23
Test.IsEqual Or_C(3, 20), 23
Test.IsEqual -16 Or 20, -12
Test.IsEqual Or_C(-16, 20), -12
Test.IsEqual 15 Or 15, 15
Test.IsEqual Or_C(15, 15), 15

'Unsigned longs test
Test.IsEqual Or_C(2606951925#, 32906, False), 2606951935#
Test.IsEqual Or_C(2264371258#, 2147516416#, False), 2264371258#
Test.IsEqual Or_C(1823284833, 2147483658#, False), 3970768491#


Set Test = Suite.Test("TestAnd")
Test.IsEqual 5 And 20, 4
Test.IsEqual And_C(5, 20), 4
Test.IsEqual -16 And 20, 16
Test.IsEqual And_C(-16, 20), 16

'Unsigned longs test
Test.IsEqual And_C(2606951925#, 32906, False), 32896
Test.IsEqual And_C(2264371258#, 2147516416#, False), 2147516416#
Test.IsEqual And_C(1823284833, 2147483658#, False), 0


Set Test = Suite.Test("TestNot")
Test.IsEqual Not 16, -17
Test.IsEqual Not_C(16), -17
Test.IsEqual Not -16, 15
Test.IsEqual Not_C(-16), 15

Test.IsEqual Not_C(2147483648#), -2147483649#
Test.IsEqual Not_C(2147483648#, False), 2147483647


End Sub

Function DecToBin_C(DecimalIn As Variant, OutputLen As Byte, Optional IsSigned As Boolean = True) As String

    If IsSigned Then
        'Signed value in, e.g. len = 16 -> -32,768 to 32,767
        MinDecVal = CDec(-2 ^ (OutputLen - 1))
        MaxDecVal = CDec(2 ^ (OutputLen - 1) - 1)
    Else
        'Unsigned value in, e.g. len = 16  -> 0 to 65535
        MinDecVal = CDec(0)
        MaxDecVal = CDec(2 ^ OutputLen - 1)
    End If
    
    DecToBin2 = ""
    DecCalc = CDec(DecimalIn)
    If DecCalc < MinDecVal Or DecCalc > MaxDecVal Then
        'Error (6) 'overflow -> error normally off, giving back an empty string, but can switch it on
        DecToBin_C = DecToBin2
        Exit Function
    End If
    
    Do While DecimalIn <> 0
        DecToBin2 = Trim$(Str$(DecCalc - 2 * Int(DecCalc / 2))) & DecToBin2
        DecCalc = Int(DecCalc / 2)
        'Escape for maximum length (negative numbers):
        If Len(DecToBin2) = OutputLen Then Exit Do
    Loop
    DecToBin_C = Right$(String$(OutputLen, "0") & DecToBin2, OutputLen)
    
End Function
Function BinToDec_C(StringIn As String, Optional IsSigned As Boolean = True) As Variant
    
    'Input assumed to be a Signed number, otherwise use IsSigned = False
    Dim StrLen As Byte
    StrLen = Len(StringIn)
    BinToDec_C = 0
    If Left(StringIn, 1) = "1" And IsSigned Then
        'negative number, signed
         For i = 1 To Len(StringIn)
            If Mid(StringIn, StrLen + 1 - i, 1) = "0" Then
                BinToDec_C = BinToDec_C + 2 ^ (i - 1)
            End If
        Next i
        BinToDec_C = -BinToDec_C - 1
    Else
        'positive number, can be signed or unsigned
        For i = 1 To Len(StringIn)
            If Mid(StringIn, StrLen + 1 - i, 1) = "1" Then
                BinToDec_C = BinToDec_C + 2 ^ (i - 1)
            End If
        Next i
    End If
    
End Function

Function LeftShift(ValIn As Variant, Shift As Byte, Optional DefaultLen As Byte = 1, Optional IsSigned As Boolean = True) As Variant
    
    '<<  Zero fill left shift - Shifts left by pushing zeros in from the right and let the leftmost bits fall off
    If DefaultLen = 1 Then
        ' DefaultLen -> will get the most appropriate value of 8 (byte), 16 (integer), 32 (long), 64 (longlong)
        DefaultLen = GetDefaultLen(ValIn, IsSigned)
    End If
    
    Dim TempStr As String
    TempStr = DecToBin_C(ValIn, DefaultLen, IsSigned)
    TempStr = Right$(TempStr & String$(Shift, "0"), DefaultLen)
    LeftShift = BinToDec_C(TempStr, IsSigned)
    
End Function

Function RightShift(ValIn As Variant, Shift As Byte, Optional DefaultLen As Byte = 1, Optional IsSigned As Boolean = True) As Variant
    
    '>>  Signed right shift  Shifts right by pushing copies of the leftmost bit in from the left, and let the rightmost bits fall off
    'Also called: Signed Right Shift [>>]
    If DefaultLen = 1 Then
        ' DefaultLen -> will get the most appropriate value of 8 (byte), 16 (integer), 32 (long), 64 (longlong)
        DefaultLen = GetDefaultLen(ValIn, IsSigned)
    End If
    
    Dim TempStr As String
    Dim FillStr As String
    TempStr = DecToBin_C(ValIn, DefaultLen, IsSigned)
    FillStr = Left(TempStr, 1)
    TempStr = Left$(String$(Shift, FillStr) & TempStr, DefaultLen)
    RightShift = BinToDec_C(TempStr, IsSigned)
    
End Function

Function RightShiftZF(ValIn As Variant, Shift As Byte, Optional DefaultLen As Byte = 1, Optional IsSigned As Boolean = True) As Variant
    
    '>>> Zero fill right shift   Shifts right by pushing zeros in from the left, and let the rightmost bits fall off
    'Also called: Unsigned Right Shift [>>>]
    If DefaultLen = 1 Then
        ' DefaultLen -> will get the most appropriate value of 8 (byte), 16 (integer), 32 (long), 64 (longlong)
        DefaultLen = GetDefaultLen(ValIn, IsSigned)
    End If
    
    Dim TempStr As String
    TempStr = DecToBin_C(ValIn, DefaultLen, IsSigned)
    TempStr = Left$(String$(Shift, "0") & TempStr, DefaultLen)
    RightShiftZF = BinToDec_C(TempStr, IsSigned)
End Function

Function HexToDec_C(hexString As String) As Variant
    'https://stackoverflow.com/questions/40213758/convert-hex-string-to-unsigned-int-vba#40217566
    'cut off "&h" if present
    If Left(hexString, 2) = "&h" Or Left(hexString, 2) = "&H" Then hexString = Mid(hexString, 3)

    'cut off leading zeros
    While Left(hexString, 1) = "0"
        hexString = Mid(hexString, 2)
    Wend
    
    If hexString = "" Then hexString = "0"
    HexToDec_C = CDec("&h" & hexString)
    'correct value for 8 digits onle
    'Debug.Print hexString, HexToDec_C
    If HexToDec_C < 0 And Len(hexString) = 8 Then
        HexToDec_C = CDec("&h1" & hexString) - 4294967296#
    'cause overflow for 16 digits
    ElseIf HexToDec_C < 0 Then
        Error (6) 'overflow
    End If

End Function
Function GetDefaultLen(ValIn As Variant, IsSigned As Boolean) As Byte

If IsSigned Then
    'Signed value in, e.g. len = 16 -> -32,768 to 32,767
    If CDec(ValIn) >= -2 ^ (8 - 1) And CDec(ValIn) <= 2 ^ (8 - 1) - 1 Then
        GetDefaultLen = 8 '8 (byte)
    ElseIf CDec(ValIn) >= -2 ^ (16 - 1) And CDec(ValIn) <= 2 ^ (16 - 1) - 1 Then
        GetDefaultLen = 16 '16 (integer)
    ElseIf CDec(ValIn) >= -2 ^ (32 - 1) And CDec(ValIn) <= 2 ^ (32 - 1) - 1 Then
        GetDefaultLen = 32 '32 (long)
    ElseIf CDec(ValIn) >= -2 ^ (64 - 1) And CDec(ValIn) <= 2 ^ (64 - 1) - 1 Then
        GetDefaultLen = 64 '64 (longlong)
    Else
        'Number too big for function, return max value that Currency can represent
        GetDefaultLen = 96
    End If
Else
    'Unsigned value in, e.g. len = 8  -> 0 to 255
    If CDec(ValIn) <= 2 ^ 8 - 1 And CDec(ValIn) >= 0 Then
        GetDefaultLen = 8 '8 (byte)
    ElseIf CDec(ValIn) <= 2 ^ 16 - 1 And CDec(ValIn) >= 0 Then
        GetDefaultLen = 16 '16 (integer)
    ElseIf CDec(ValIn) <= 2 ^ 32 - 1 And CDec(ValIn) >= 0 Then
        GetDefaultLen = 32 '32 (long)
    ElseIf CDec(ValIn) <= 2 ^ 64 - 1 And CDec(ValIn) >= 0 Then
        GetDefaultLen = 64 '64 (longlong)
    Else
        'Number too big for function, return max value that Currency can represent
        GetDefaultLen = 96
    End If
End If

End Function

Function Not_C(ValIn1 As Variant, Optional IsSigned As Boolean = True) As Variant
    
    Dim s3 As String
    Dim s1len As Byte
    d1 = CDec(ValIn1)
    
    UseDefault = True
    If IsSigned = True Then
        If d1 < -2 ^ (32 - 1) Or d1 > 2 ^ (32 - 1) - 1 Then UseDefault = False
    Else
        UseDefault = False
    End If
    
    If UseDefault Then
        Not_C = Not ValIn1
    Else
        'Check size and sign
        s1len = GetDefaultLen(d1, IsSigned)
        s1 = DecToBin_C(d1, s1len, IsSigned)
        s3 = ""
        For C = 1 To s1len
            If Mid(s1, C, 1) = "1" Then
                s3 = s3 & "0"
            Else
                s3 = s3 & "1"
            End If
        Next C
        Not_C = BinToDec_C(s3, IsSigned)
    End If
    
End Function
Function And_C(ValIn1 As Variant, ValIn2 As Variant, Optional IsSigned As Boolean = True) As Variant
    And_C = OrAndXor_C("AND", ValIn1, ValIn2, IsSigned)
End Function
Function Or_C(ValIn1 As Variant, ValIn2 As Variant, Optional IsSigned As Boolean = True) As Variant
    Or_C = OrAndXor_C("OR", ValIn1, ValIn2, IsSigned)
End Function
Function Xor_C(ValIn1 As Variant, ValIn2 As Variant, Optional IsSigned As Boolean = True) As Variant
    Xor_C = OrAndXor_C("XOR", ValIn1, ValIn2, IsSigned)
End Function
Function OrAndXor_C(Func As String, ValIn1 As Variant, ValIn2 As Variant, Optional IsSigned As Boolean = True) As Variant

    Dim s3 As String
    Dim maxlen As Byte
    d1 = CDec(ValIn1)
    d2 = CDec(ValIn2)
    Func = LCase(Func)
    
    UseDefault = True
    If IsSigned = True Then
        If d1 < -2 ^ (32 - 1) Or d1 > 2 ^ (32 - 1) - 1 Then UseDefault = False
        If d2 < -2 ^ (32 - 1) Or d2 > 2 ^ (32 - 1) - 1 Then UseDefault = False
    Else
        UseDefault = False
    End If
    
    If UseDefault Then
        If Func = "xor" Then
            OrAndXor_C = d1 Xor d2
        ElseIf Func = "or" Then
            OrAndXor_C = d1 Or d2
        ElseIf Func = "and" Then
            OrAndXor_C = d1 And d2
        Else
            OrAndXor_C = False
        End If
    Else
        If IsSigned Then
            'Too big for a 32 bit long, go for 64 bit
            s1 = DecToBin_C(d1, 64)
            s2 = DecToBin_C(d2, 64)
            s3 = ""
            For C = 1 To 64
                If Func = "xor" Then
                    If Mid(s1, C, 1) = Mid(s2, C, 1) Then
                        s3 = s3 & "0"
                    Else
                        s3 = s3 & "1"
                    End If
                ElseIf Func = "or" Then
                    If Mid(s1, C, 1) = 1 Or Mid(s2, C, 1) = 1 Then
                        s3 = s3 & "1"
                    Else
                        s3 = s3 & "0"
                    End If
                ElseIf Func = "and" Then
                    If Mid(s1, C, 1) = 1 And Mid(s2, C, 1) = 1 Then
                        s3 = s3 & "1"
                    Else
                        s3 = s3 & "0"
                    End If
                End If
            Next C
            OrAndXor_C = BinToDec_C(s3)
        Else
            'Treat as unsigned
            s1len = GetDefaultLen(d1, False)
            s2len = GetDefaultLen(d2, False)
            
            If s1len > s2len Then maxlen = s1len Else maxlen = s2len
            
            s1 = DecToBin_C(d1, maxlen, False)
            s2 = DecToBin_C(d2, maxlen, False)
            s3 = ""
            
            For C = 1 To maxlen
                If Func = "xor" Then
                    If Mid(s1, C, 1) = Mid(s2, C, 1) Then
                        s3 = s3 & "0"
                    Else
                        s3 = s3 & "1"
                    End If
                ElseIf Func = "or" Then
                    If Mid(s1, C, 1) = 1 Or Mid(s2, C, 1) = 1 Then
                        s3 = s3 & "1"
                    Else
                        s3 = s3 & "0"
                    End If
                ElseIf Func = "and" Then
                    If Mid(s1, C, 1) = 1 And Mid(s2, C, 1) = 1 Then
                        s3 = s3 & "1"
                    Else
                        s3 = s3 & "0"
                    End If
                End If
            Next C
            OrAndXor_C = BinToDec_C(s3, False)
        
        End If
    End If

End Function

'to work with
'Integer 2 bytes (16 bits)   -32,768 to 32,767   0   vbInteger
'Long (long integer) 4 bytes (32 bits)   -2,147,483,648 to 2,147,483,647
'LongLong (LongLong integer) 8 bytes (64 bits)   -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
'Set up to work with LONG variables
'Treat the data as unsigned LONGs
'Work with DECIMAL, can hold both Long, ULong (unsigned long), LongLong and ULongLong (64 bits)
'Unsigned Char 0b11111111 (0xFF in hex) = 255 in decimal, (128+64+32+16+8+4+2+1 = 255)
'Signed Char 0b11111111 (0xFF in hex) = -127 in decimal, (-1 * (64+32+16+8+4+2+1) = - 127)
'NR, SHIFT 3 LEFT, SHIFT 3 RIGHT
'CASE: 0  0  0
'CASE: 100000  800000  12500
'CASE: -100000  -800000  536858412 (-12500 is Signed right shift)
'CASE: 1747148294  1092284464  218393536
'CASE: -195521891  -1564175128  512430675
'CASE: -2147483648  0  268435456


'Sub TestA()
'
''2d array
'Dim RCs
'RCs = Array("0000000000000001", "0000000000008082", "800000000000808a", "8000000080008000", "000000000000808b", "0000000080000001", _
'            "8000000080008081", "8000000000008009", "000000000000008a", "0000000000000088", "0000000080008009", "000000008000000a", _
'            "000000008000808b", "800000000000008b", "8000000000008089", "8000000000008003", "8000000000008002", "8000000000000080", _
'            "000000000000800a", "800000008000000a", "8000000080008081", "8000000000008080", "0000000080000001", "8000000080008008")
'Dim RC(0 To 23, 0 To 1) As Currency
'
'For R = 0 To UBound(RCs)
'    Debug.Print "LOOP 1: ", R, RCs(R)
'    RC(R, 0) = HexToDec_C(Right(RCs(R), 8))
'    RC(R, 1) = HexToDec_C(Left(RCs(R), 8))
'    'Debug.Print "hi " & RC(r, 1) & "   lo " & RC(r, 0)
'Next R
'
'End Sub
'Sub TestShift()
'Dim v As Long
'Dim s As String
'v = -100000
's = DecToBin32(v)
't = Bin32ToDec(s)
'Debug.Print v, s, t
'
'v = -9
'Debug.Print v
'Debug.Print LeftShift(v, 2)
'Debug.Print RightShift(v, 2)
'Debug.Print RightShiftZF(v, 2)
'
'
''Operator    Name    Description
''&   AND Sets each bit to 1 if both bits are 1
''|   OR  Sets each bit to 1 if one of two bits is 1
''^   XOR Sets each bit to 1 if only one of two bits is 1
''~   NOT Inverts all the bits
''<<  Zero fill left shift    Shifts left by pushing zeros in from the right and let the leftmost bits fall off
''>>  Signed right shift  Shifts right by pushing copies of the leftmost bit in from the left, and let the rightmost bits fall off
''>>> Zero fill right shift   Shifts right by pushing zeros in from the left, and let the rightmost bits fall off
'
'
''n = 1
''m = 32 - n
''Debug.Print 0, ShiftLeft(0, 3), ShiftRight(0, 3), shr(0, 3)
''Debug.Print 100000, ShiftLeft(100000, 3), ShiftRight(100000, 3), shr(100000, 3)
''Debug.Print -100000, ShiftLeft(-100000, 3), ShiftRight(-100000, 3), shr(-100000, 3)
''Debug.Print 1747148294, ShiftLeft(1747148294, 3), ShiftRight(1747148294, 3), shr(1747148294, 3)
''Debug.Print -195521891, ShiftLeft(-195521891, 3), ShiftRight(-195521891, 3), shr(-195521891, 3)
''Debug.Print -2147483648#, ShiftLeft(-2147483648#, 3), ShiftRight(-2147483648#, 3), shr(-2147483648#, 3)
'
''SHIFT 3 left & right:
''CASE: 0  0  0
''CASE: 1747148294  1092284464  218393536
''CASE: -195521891  -1564175128  512430675
''CASE: -2147483648  0  268435456
'
''NR, SHIFT 3 LEFT, SHIFT 3 RIGHT
''CASE: 0  0  0
''CASE: 100000  800000  12500
''CASE: -100000  -800000  536858412 (-12500 is Signed right shift)
''CASE: 1747148294  1092284464  218393536
''CASE: -195521891  -1564175128  512430675
''CASE: -2147483648  0  268435456
'A = 2147483648#
'   '2147516545
'
''-9 : 11111111111111111111111111110111
''+9 : 00000000000000000000000000001001
'
''32 bit numbers: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Bitwise_Operators
'
'End Sub






'Function ShiftBit(s As String, n As Long) As Long
'    'https://www.pcreview.co.uk/threads/bitwise-shift-in-vba.3883630/
'    Dim L As Long
'    L = Asc(s)
'    ShiftBit = L * (2 ^ n)
'End Function

'Function ShiftLeft(ByVal Value As Long, ByVal Shift As Byte) As Long
'    'http://www.excely.com/excel-vba/bit-shifting-function.shtml
'    ShiftLeft = Value
'    If Shift > 0 Then
'        Dim i As Byte
'        Dim m As Long
'        For i = 1 To Shift
'            m = ShiftLeft And &H40000000
'            ShiftLeft = (ShiftLeft And &H3FFFFFFF) * 2
'            If m <> 0 Then
'                ShiftLeft = ShiftLeft Or &H80000000
'            End If
'        Next i
'    End If
'End Function
'
'Function ShiftRight(ByVal Value As Currency, ByVal Shift As Byte) As Currency
'    'http://www.excely.com/excel-vba/bit-shifting-function.shtml
'    Dim i As Byte
'    ShiftRight = Value
'    If Shift > 0 Then
'        ShiftRight = Int((ShiftRight / (2 ^ Shift)) And Not &H80000000)
'    End If
'End Function

'Sub testshift()
'
'Debug.Print ShiftBit("h", 0)
'Debug.Print ShiftBit("e", 8)
'Debug.Print ShiftBit("l", 16)
'Debug.Print ShiftBit("l", 24)
'
'Debug.Print ShiftLeft("h", 0)
'Debug.Print ShiftLeft("e", 8)
'Debug.Print ShiftLeft("l", 16)
'Debug.Print ShiftLeft("l", 24)
'
'End Sub

