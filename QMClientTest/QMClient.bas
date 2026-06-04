Attribute VB_Name = "QMClient"
Declare Function QMConnect Lib "QMClient.dll" ( _
    ByVal hostPtr As Long, _
    ByVal port As Integer, _
    ByVal usernamePtr As Long, _
    ByVal passwordPtr As Long, _
    ByVal accountPtr As Long) As Integer

Declare Function QMConnected Lib "QMClient.dll" () As Integer

Declare Sub QMDisconnect Lib "QMClient.dll" ()

Declare Sub QMCall Lib "QMClient.dll" ( _
    ByVal subrnamePtr As Long, _
    ByVal argc As Integer, _
    ByVal a1Ptr As Long, ByVal a2Ptr As Long, ByVal a3Ptr As Long, ByVal a4Ptr As Long, _
    ByVal a5Ptr As Long, ByVal a6Ptr As Long, ByVal a7Ptr As Long, ByVal a8Ptr As Long, _
    ByVal a9Ptr As Long, ByVal a10Ptr As Long, ByVal a11Ptr As Long, ByVal a12Ptr As Long, _
    ByVal a13Ptr As Long, ByVal a14Ptr As Long, ByVal a15Ptr As Long, ByVal a16Ptr As Long, _
    ByVal a17Ptr As Long, ByVal a18Ptr As Long, ByVal a19Ptr As Long, ByVal a20Ptr As Long)
    
Declare Function QMOpen Lib "QMClient.dll" ( _
    ByVal filenamePtr As Long, _
    ByVal mode As Integer) As Integer  ' returns file number (fno) as Integer
    
Declare Function QMRead Lib "QMClient.dll" ( _
    ByVal fno As Integer, _
    ByVal idPtr As Long, _
    ByRef err As Integer, _
    ByVal mode As Integer) As Long  ' returns BSTR pointer as Long
    
Declare Function QMWrite Lib "QMClient.dll" ( _
    ByVal fno As Integer, _
    ByVal idPtr As Long, _
    ByVal dataPtr As Long) As Integer
    
Declare Function QMChange Lib "QMClient.dll" ( _
    ByVal srcPtr As Long, _
    ByVal oldPtr As Long, _
    ByVal newPtr As Long, _
    ByRef occ As Long, _
    ByRef first As Long) As Long  ' returns BSTR pointer as Long
    
Declare Function QMStatus Lib "QMClient.dll" ( _
    ByVal something As Long) As Long

'-- Helper Win32 / OLE API declares for conversion and memory ops:

Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)

Declare Function SysStringByteLen Lib "oleaut32.dll" (ByVal bstr As Long) As Long

Declare Sub SysFreeString Lib "oleaut32.dll" (ByVal bstr As Long)

' Convert returned BSTR pointer to VB String
Function BSTRPtrToString(bstrPtr As Long) As String
    ' Notes
    ' * Always use VarPtr(buf(0)) (and pass it ByVal) to give RtlMoveMemory the address.
    ' * Check n>0 before creating/using buf(0) to avoid errors on zero-length strings.
    ' * Alternative declare (CopyMemory with As Any) exists, but the ByVal Long form is simpler and avoids type mismatch.
    If bstrPtr = 0 Then Exit Function
    Dim n As Long: n = SysStringByteLen(bstrPtr)
    If n <= 0 Then
        SysFreeString bstrPtr
        Exit Function
    End If
    Dim buf() As Byte
    ReDim buf(0 To n - 1) As Byte ' Copy into byte array using the array element address
    RtlMoveMemory ByVal VarPtr(buf(0)), ByVal bstrPtr, n
    BSTRPtrToString = StrConv(buf, vbUnicode)
    SysFreeString bstrPtr
End Function

'-- Example usage:

'--
'   Dim rc As Integer
'   rc = QMConnect(StrPtr("host.example"), 5010, StrPtr("user"), StrPtr("pass"), StrPtr("account"))
'   If rc <> 0 Then ' connected
'   End If

'--
'   Dim err As Integer
'   Dim bstrPtr As Long
'   bstrPtr = QMRead(1, StrPtr("ID123"), err, 0)
'   If bstrPtr <> 0 Then
'       Dim text As String
'       text = BSTRPtrToString(bstrPtr) ' use text
'   End If

Public Function PtrFor(s As String) As Long
    If Len(s) = 0 Then
        PtrFor = 0
    Else
        PtrFor = StrPtr(s)
    End If
End Function


