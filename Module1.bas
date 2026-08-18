Attribute VB_Name = "Module1"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function ByteToStr(bArray() As Byte) As String
    Dim lPntr As Long
    Dim bTmp() As Byte
    On Error GoTo ByteErr
    ReDim bTmp(UBound(bArray) * 2 + 1)
    For lPntr = 0 To UBound(bArray)
        bTmp(lPntr * 2) = bArray(lPntr)
    Next lPntr
    Let ByteToStr = bTmp
    Exit Function
ByteErr:
    ByteToStr = ""
End Function

Public Function ByteToUni(bArray() As Byte) As String
    ByteToUni = bArray
End Function

Public Sub DebugPrintByte(sDescr As String, bArray() As Byte)
    Dim lPtr As Long
    Debug.Print sDescr & ":"
    If GetbSize(bArray) = 0 Then Exit Sub
    For lPtr = 0 To UBound(bArray)
        Debug.Print right$("0" & Hex$(bArray(lPtr)), 2) & " ";
        If (lPtr + 1) Mod 16 = 0 Then Debug.Print
    Next lPtr
    Debug.Print
End Sub

Public Function GetbSize(bArray() As Byte) As Long
    On Error GoTo GetSizeErr
    GetbSize = UBound(bArray) + 1
    Exit Function
GetSizeErr:
    GetbSize = 0
End Function

Public Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function

Public Sub DebugPrintString(sDescr As String, strToPrint As String)
    Dim lPtr As Long
    Dim sSep As String * 1
    Debug.Print sDescr & ":"
    For lPtr = 0 To LenB(strToPrint) - 1
        Debug.Print right$("0" & Hex$(PeekB(StrPtr(strToPrint) + lPtr)), 2) & " ";
        If (lPtr + 2) Mod 16 = 0 Then Debug.Print
    Next lPtr
    Debug.Print
End Sub

Public Function StrToByte(strInput As String) As Byte()
    Dim lPntr As Long
    Dim bTmp() As Byte
    Dim bArray() As Byte
    If Len(strInput) = 0 Then Exit Function
    ReDim bTmp(LenB(strInput) - 1) 'Memory length
    ReDim bArray(Len(strInput) - 1) 'String length
    CopyMemory bTmp(0), ByVal StrPtr(strInput), LenB(strInput)
    'Examine every second byte
    For lPntr = 0 To UBound(bArray)
        If bTmp(lPntr * 2 + 1) > 0 Then
            'bArray(lPntr) = Asc(Mid$(strInput, lPntr + 1, 1))
            StrToByte = bTmp
            Exit Function
        Else
            bArray(lPntr) = bTmp(lPntr * 2)
        End If
    Next lPntr
    StrToByte = bArray
End Function

Public Function UniToByte(strInput As String) As Byte()
    UniToByte = strInput
End Function

