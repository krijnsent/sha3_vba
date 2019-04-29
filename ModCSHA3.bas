Attribute VB_Name = "ModCSHA3"
'https://keccak.team/software.html
'PRIMARY SOURCE: https://www.movable-type.co.uk/scripts/sha3.html#src-code
'SECONDARY SOURCE: https://github.com/mjosaarinen/tiny_sha3
'http://castoro.nl/sha3_veness.html

Sub TestSHA3Functions()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModSHA3Functions"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim SHATestTxt As String
Dim Test As TestCase
Set Test = Suite.Test("SHA3_basics")
'Simple test
SHATestTxt = "hello"
Test.IsEqual Hash224(SHATestTxt), "b87f88c72702fff1748e58b87e9141a42c0dbedc29a78cb0d4a5cd81"
SHATestTxt = "hello"
Test.IsEqual Hash256(SHATestTxt), "3338be694f50c5f338814986cdf0686453a888b84f424d792af4b9202398f392"
SHATestTxt = "hello"
Test.IsEqual Hash384(SHATestTxt), "720aea11019ef06440fbf05d87aa24680a2153df3907b23631e7177ce620fa1330ff07c0fddee54699a4c3ee0ee9d887"
SHATestTxt = "hello"
Test.IsEqual Hash512(SHATestTxt), "75d527c368f2efe848ecf6b073a36767800805e9eef2b1857d5f984f036eb6df891d75f72d9b154518c1cd58835286d1da9a38deba3de98b5a53e5ed78a84976"

Set Test = Suite.Test("SHA3_longer_text")
SHATestTxt = "Let me tell you something about building this thing in VBA... It's a horror and took way too long! Debug.print and console.log have been my biggest friends on this journey. Do not try this at home :P... Koen"
Test.IsEqual Hash224(SHATestTxt), "574cd126b7595b85cedf42337d6d6b49776455249e8c4bbd15c7639c"
SHATestTxt = "Let me tell you something about building this thing in VBA... It's a horror and took way too long! Debug.print and console.log have been my biggest friends on this journey. Do not try this at home :P... Koen"
Test.IsEqual Hash256(SHATestTxt), "cb26f5a71cbc74a11b5662236d03ac615168f01ed41398db5d9963ef36ab9b8a"
SHATestTxt = "Let me tell you something about building this thing in VBA... It's a horror and took way too long! Debug.print and console.log have been my biggest friends on this journey. Do not try this at home :P... Koen"
Test.IsEqual Hash384(SHATestTxt), "c3d8e28f1ea3b8b0c5c0aec9a2ae71c03b1cc41e4edfafdad3d91728d573e052f581efdf862ce9f66acf669f5e80ec50"
SHATestTxt = "Let me tell you something about building this thing in VBA... It's a horror and took way too long! Debug.print and console.log have been my biggest friends on this journey. Do not try this at home :P... Koen"
Test.IsEqual Hash512(SHATestTxt), "a49c893e289c02573a6da9f202532e1f405440c9c16318e41309f43f330753bf9f70983b1729bca23ec9f0315b2f62ea84c634b8d0081966fbb002f94c6664db"

Set Test = Suite.Test("SHA3_special_chars")
SHATestTxt = "1234567890-=!@#$%^&*()_+[]{};:<>,./?"
Test.IsEqual Hash256(SHATestTxt), "8fcb737415368ddd535c7b1b19b53c42bc2eca7cb7d4f7c3ec28b601da76b63d"

End Sub

Function Hash224(msg As String, Optional opt As Dictionary) As String

'Generates 224-bit SHA-3 / Keccak hash of message.
'String msg - String to be hashed (Unicode-safe).
'Dictionary options - padding: sha-3 / keccak; msgFormat: string / hex; outFormat: hex / hex-b / hex-w.
Hash224 = Keccak1600(1152, 448, msg, opt)

End Function
Function Hash256(msg As String, Optional opt As Dictionary) As String

'Generates 224-bit SHA-3 / Keccak hash of message.
Hash256 = Keccak1600(1088, 512, msg, opt)

End Function

Function Hash384(msg As String, Optional opt As Dictionary) As String

'Generates 384-bit SHA-3 / Keccak hash of message.
Hash384 = Keccak1600(832, 768, msg, opt)

End Function
Function Hash512(msg As String, Optional opt As Dictionary) As String

'Generates 512-bit SHA-3 / Keccak hash of message.
Hash512 = Keccak1600(576, 1024, msg, opt)

End Function
Function Keccak1600(R As Integer, C As Integer, msg As String, Optional opt As Dictionary) As String
    
'Generates SHA-3 / Keccak hash of message M.
'Integer r - Bitrate 'r' (b-c)
'Integer c - Capacity 'c' (b-r), md length × 2
'String msg - Message
'Dictionary options - padding: sha-3 / keccak; msgFormat: string / hex; outFormat: hex / hex-b / hex-w.
'{string} Hash as hex-encoded string.
    
    
'const defaults = { padding: 'sha-3', msgFormat: 'string', outFormat: 'hex' };
Set OptDefaults = New Scripting.Dictionary
OptDefaults.Add "padding", "sha-3"
OptDefaults.Add "msgFormat", "string"
OptDefaults.Add "outFormat", "hex"
    
If opt Is Nothing Then Set opt = New Scripting.Dictionary
For Each k In OptDefaults.Keys
    If Not opt.Exists(k) Then
        opt.Add k, OptDefaults(k)
    End If
Next k
    
MsgLen = C / 2
' message digest output length in bits
    
'
If opt("msgFormat") = "hex-bytes" Then
    'NOT IMPLEMENTED YET, hexBytesToString(M)
    'msg = StrConv(msg, vbUnicode)
Else
    'utf8Encode(M)
    'msg = StrConv(msg, vbUnicode)
End If

'2d array
Dim state(0 To 4, 0 To 4, 0 To 1) As Currency
Dim squeezeState(0 To 4, 0 To 4) As String
' last dimension: 0 = lo, 1 = hi
' * Keccak state is a 5 × 5 x w array of bits (w=64 for keccak-f[1600] / SHA-3).
' * Here, it is implemented as a 5 × 5 array of Long. The first subscript (x) defines the
' * sheet, the second (y) defines the plane, together they define a lane. Slices, columns,
' * and individual bits are obtained by bit operations on the hi,lo components of the Long
' * representing the lane.

q = (R / 8) - Len(msg) Mod (R / 8)
If q = 1 Then
    If opt("padding") = "keccak" Then
        msg = msg & Chr$(129)
    Else
        msg = msg & Chr$(134)
    End If
Else
    If opt("padding") = "keccak" Then
        msg = msg & Chr$(1)
    Else
        msg = msg & Chr$(6)
    End If
    msg = msg & String(q - 2, Chr$(0))
    msg = msg & Chr$(128)
End If

'Debug.Print "q", q, Len(msg), msg,

w = 64  'for keccak-f[1600]
blocksize = R / w * 8

'Debug.Print w, blocksize

i = 0
Do While i < Len(msg)
    j = 0
    Do While j < R / w
        lo = LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 0 + 1, 1))), 0, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 1 + 1, 1))), 8, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 2 + 1, 1))), 16, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 3 + 1, 1))), 24, 32)
        hi = LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 4 + 1, 1))), 0, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 5 + 1, 1))), 8, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 6 + 1, 1))), 16, 32) + _
                LeftShift(CLng(Asc(Mid(msg, i + j * 8 + 7 + 1, 1))), 24, 32)
        x = j Mod 5
        y = Int(j / 5)
        
        'Debug.Print "x,y lo,hi ", i & "," & j & "  " & lo & "," & hi
        state(x, y, 0) = state(x, y, 0) Xor lo
        state(x, y, 1) = state(x, y, 1) Xor hi
        j = j + 1
    Loop

    newstate = keccak_f_1600(state)
    
    i = i + blocksize
Loop

'Squeeze state
For i = 0 To 4
    For j = 0 To 4
        v1 = state(i, j, 0)
        v2 = state(i, j, 1)
        If v1 >= 2 ^ (32 - 1) Then v1 = v1 - 2 ^ (32)
        If v2 >= 2 ^ (32 - 1) Then v2 = v2 - 2 ^ (32)
        s1 = Hex(v1)
        s2 = Hex(v2)
        If Len(s1) < 8 Then s1 = String$(8 - Len(s1), "0") & s1
        If Len(s2) < 8 Then s2 = String$(8 - Len(s2), "0") & s2
        
        squeezeState(i, j) = LCase(s2 & s1)
        'Debug.Print i, j, squeezeState(i, j)
    Next j
Next i

ResStr = ""
For j = 0 To 4
    For i = 0 To 4
        For k = 8 To 1 Step -1
            ResStr = ResStr & Mid(squeezeState(i, j), 2 * k - 1, 2)
        Next k
        'Debug.Print ResStr
    Next i
Next j

Keccak1600 = Left(ResStr, MsgLen / 4)

'// if required, group message digest into bytes or words
'if (opt.outFormat == 'hex-b') md = md.match(/.{2}/g).join(' ');
'if (opt.outFormat == 'hex-w') md = md.match(/.{8,16}/g).join(' ');

'Debug.Print "END HERE!"
'550b320103b1f401"
'550b32013b1f401
'b87f88c72702fff1748e58b87e9141a42c0dbedc29a78cb0d4a5cd81a96abded
'b87f88c72702fff1748e58b87e9141a42c0dbedc29a78cb0d4a5cd81a96abded52f214ef4fb788ba

End Function

Function keccak_f_1600(StateIn)

nRounds = 24

'2d array
Dim RCs
RCs = Array("0000000000000001", "0000000000008082", "800000000000808a", "8000000080008000", "000000000000808b", "0000000080000001", _
            "8000000080008081", "8000000000008009", "000000000000008a", "0000000000000088", "0000000080008009", "000000008000000a", _
            "000000008000808b", "800000000000008b", "8000000000008089", "8000000000008003", "8000000000008002", "8000000000000080", _
            "000000000000800a", "800000008000000a", "8000000080008081", "8000000000008080", "0000000080000001", "8000000080008008")
Dim RC(0 To 23, 0 To 1) As Currency

For R = 0 To UBound(RCs)
    RC(R, 0) = HexToDec_C(Right(RCs(R), 8))
    RC(R, 1) = HexToDec_C(Left(RCs(R), 8))
    'Put data back into Long range, as shifts are binary
    If RC(R, 0) >= 2 ^ (32 - 1) Then RC(R, 0) = RC(R, 0) - 2 ^ (32)
    If RC(R, 1) >= 2 ^ (32 - 1) Then RC(R, 1) = RC(R, 1) - 2 ^ (32)
    'Debug.Print "hi " & RC(R, 1) & "   lo " & RC(R, 0)
Next R

'// Keccak-f permutations
For R = 0 To nRounds - 1
    'Debug.Print "r:" & R
    'Debug.Print "Keccak 2.3.2"
    'Debug.Print StateIn(0, 0, 0), StateIn(0, 0, 1)

    Dim C(0 To 4, 0 To 1) As Currency
    For x = 0 To 4
        C(x, 0) = StateIn(x, 0, 0)
        C(x, 1) = StateIn(x, 0, 1)
        For y = 1 To 4
            'Debug.Print "xy chi " & x & y & "  " & C(x, 1)
            'Debug.Print "xy clo " & x & y & "  " & C(x, 0)
            C(x, 1) = Xor_C(C(x, 1), StateIn(x, y, 1))
            C(x, 0) = Xor_C(C(x, 0), StateIn(x, y, 0))
        Next y
    Next x
    
    'Debug.Print "Keccak 2.3.2 bis"
    'Debug.Print StateIn(0, 0, 0), StateIn(0, 0, 1)
    
    For x = 0 To 4
        'Debug.Print "D hi- " & x & "  " & C((x + 4) Mod 5, 1)
        'Debug.Print "D lo- " & x & "  " & C((x + 4) Mod 5, 0)
        Dim Rt(0 To 1) As Currency
        Rt(0) = C((x + 1) Mod 5, 0)
        Rt(1) = C((x + 1) Mod 5, 1)
        Rr = rotl(Rt, 1)
        'Debug.Print "D rot hi- " & x & "  " & Rr(1)
        'Debug.Print "D rot lo- " & x & "  " & Rr(0)
        
        hi = Xor_C(C((x + 4) Mod 5, 1), Rr(1))
        lo = Xor_C(C((x + 4) Mod 5, 0), Rr(0))
        Dim D(0 To 4, 0 To 1) As Currency
        D(x, 1) = hi
        D(x, 0) = lo
        For y = 0 To 4
            StateIn(x, y, 1) = Xor_C(StateIn(x, y, 1), D(x, 1))
            StateIn(x, y, 0) = Xor_C(StateIn(x, y, 0), D(x, 0))
        Next y
    Next x
    
    'Debug.Print "Keccak 2.3.4"
    'Debug.Print StateIn(0, 0, 0), StateIn(0, 0, 1)
    
    xa = 1
    ya = 0
    Dim tmp(0 To 1) As Currency
    Dim cur(0 To 1) As Currency
    'ReDim Rt(0 To 1) As Long
    cur(0) = StateIn(xa, ya, 0)
    cur(1) = StateIn(xa, ya, 1)
    For t = 0 To 23
        xb = ya
        yb = (2 * xa + 3 * ya) Mod 5
        'Debug.Print t, xb, yb
        tmp(0) = StateIn(xb, yb, 0)
        tmp(1) = StateIn(xb, yb, 1)
        
        Rr = rotl(cur, ((t + 1) * (t + 2) / 2) Mod 64)
        StateIn(xb, yb, 0) = Rr(0)
        StateIn(xb, yb, 1) = Rr(1)
        
        cur(0) = tmp(0)
        cur(1) = tmp(1)
        
        xa = xb
        ya = yb
    Next t
    
    
    'Debug.Print "Keccak 2.3.1"
    'Debug.Print StateIn(0, 0, 0), StateIn(0, 0, 1)
    
    For y = 0 To 4
        Erase C
        For x = 0 To 4
            C(x, 0) = StateIn(x, y, 0)
            C(x, 1) = StateIn(x, y, 1)
        Next x
        For x = 0 To 4
            StateIn(x, y, 1) = RightShiftZF(Xor_C(C(x, 1), And_C(Not_C(C((x + 1) Mod 5, 1)), C((x + 2) Mod 5, 1))), 0)
            StateIn(x, y, 0) = RightShiftZF(Xor_C(C(x, 0), And_C(Not_C(C((x + 1) Mod 5, 0)), C((x + 2) Mod 5, 0))), 0)
            'StateIn(x, y, 1) = RightShiftZF(C(x, 1) Xor ((Not C((x + 1) Mod 5, 1) And C((x + 2) Mod 5, 1))), 0)
            'StateIn(x, y, 0) = RightShiftZF(C(x, 0) Xor ((Not C((x + 1) Mod 5, 0) And C((x + 2) Mod 5, 0))), 0)
        Next x
    Next y

    'Debug.Print "Keccak 2.3.5"
    'Debug.Print StateIn(0, 0, 0), StateIn(0, 0, 1)

    'Debug.Print "a00-lo1:", StateIn(0, 0, 0), DecToBin_C(StateIn(0, 0, 0), 32)
    'Debug.Print "RCr-lo1:", RC(R, 0), DecToBin_C(StateIn(0, 0, 0), 32)
    
    StateIn(0, 0, 1) = RightShiftZF(Xor_C(StateIn(0, 0, 1), RC(R, 1)), 0)
    StateIn(0, 0, 0) = RightShiftZF(Xor_C(StateIn(0, 0, 0), RC(R, 0)), 0)

    'Debug.Print "a00-lo2:", StateIn(0, 0, 0), DecToBin_C(StateIn(0, 0, 0), 32)
    
Next R

End Function


Function rotl(ObjIn() As Currency, n As Byte) As Currency()
    
    'Debug.Print "ROTL data: ", ObjIn(0), ObjIn(1), n
    
    Dim m As Byte
    'Rotate left
    Dim R(0 To 1) As Currency
    If n < 32 Then
        m = 32 - n
        lo_1 = LeftShift(ObjIn(0), n, 32)
        lo_2 = RightShiftZF(ObjIn(1), m, 32)
        hi_1 = LeftShift(ObjIn(1), n, 32)
        hi_2 = RightShiftZF(ObjIn(0), m, 32)
        
        lo = lo_1 Or lo_2
        hi = hi_1 Or hi_2
'       const lo = this.lo<<n | this.hi>>>m;
'       const hi = this.hi<<n | this.lo>>>m;
        R(0) = lo
        R(1) = hi
    ElseIf n = 32 Then
        R(0) = ObjIn(0)
        R(1) = ObjIn(1)
    ElseIf n > 32 Then
        n = n - 32
        m = 32 - n
        lo_1 = LeftShift(ObjIn(1), n, 32)
        lo_2 = RightShiftZF(ObjIn(0), m, 32)
        hi_1 = LeftShift(ObjIn(0), n, 32)
        hi_2 = RightShiftZF(ObjIn(1), m, 32)
        lo = lo_1 Or lo_2
        hi = hi_1 Or hi_2
'       const lo = this.hi<<n | this.lo>>>m;
'       const hi = this.lo<<n | this.hi>>>m;
        R(0) = lo
        R(1) = hi
    End If
    rotl = R()
    
End Function

