Attribute VB_Name = "Module1"
Public Const SUBLANG_ENGLISH_AUS = &H3
Public Const SUBLANG_ENGLISH_CAN = &H4
Public Const SUBLANG_ENGLISH_EIRE = &H6
Public Const SUBLANG_ENGLISH_NZ = &H5
Public Const SUBLANG_ENGLISH_UK = &H2
Public Const SUBLANG_ENGLISH_US = &H1

Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Const LANG_BULGARIAN = &H2
Public Const LANG_CHINESE = &H4
Public Const LANG_CROATIAN = &H1A
Public Const LANG_CZECH = &H5
Public Const LANG_DANISH = &H6
Public Const LANG_DUTCH = &H13
Public Const LANG_ENGLISH = &H9
Public Const LANG_FINNISH = &HB
Public Const LANG_FRENCH = &HC
Public Const LANG_GERMAN = &H7
Public Const LANG_GREEK = &H8
Public Const LANG_HUNGARIAN = &HE
Public Const LANG_ICELANDIC = &HF
Public Const LANG_ITALIAN = &H10
Public Const LANG_JAPANESE = &H11
Public Const LANG_KOREAN = &H12
Public Const LANG_NEUTRAL = &H0
Public Const LANG_NORWEGIAN = &H14
Public Const LANG_POLISH = &H15
Public Const LANG_PORTUGUESE = &H16
Public Const LANG_ROMANIAN = &H18
Public Const LANG_RUSSIAN = &H19
Public Const LANG_SLOVAK = &H1B
Public Const LANG_SLOVENIAN = &H24
Public Const LANG_SPANISH = &HA
Public Const LANG_SWEDISH = &H1D
Public Const LANG_TURKISH = &H1F

Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As _
   Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type

Public Type BitmapExtract
    BitmapEnd As String
    BitmapStart As String
    BitmapHeader As String
End Type
Public Const RT_BITMAP = 2&
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As Size) As Long
Public Type Size
        cx As Long
        cy As Long
End Type

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Public Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Public Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

'Dim BitmapEnd As String
'Dim BitmapStart As String

Function LoadBitmapA(FileName As String, BitmapName As String, Pic As PictureBox)
Dim hInstance, lReturned, lRes, BMnewDC, newHDC, hHandle As Long
Dim lBuffer As String * 255
Dim MySize As Size

'loadimage
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
'lRes = FindResource(hHandle, BitmapName, RT_BITMAP)
'MsgBox GetLastError
hObject = LoadResource(hHandle, lRes)
lReturned = LoadBitmap(hHandle, BitmapName)
'MsgBox "LoadBitmap returned " & lReturned
'MsgBox GetBitmapDimensionEx(lReturned, MySize)

'Pic.Height = MySize.cx '* 150
'Pic.Width = MySize.cy '* 15
'Debug.Print "cY = " & MySize.cy
'Debug.Print "cX = " & MySize.cx
Pic.Cls
'Pic.Visible = False
x = SelectObject(Pic.hdc, lReturned)
newdc = CreateCompatibleDC(Pic.hdc)
y = SelectObject(newdc, lReturned)
'Pic.Visible = True

FreeLibrary hHandle


'hHandle = LoadBitmap(hInstance, "#159")
End Function


Function GetBitmapsA(FileName As String)
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
For i = 1 To 9999
    lReturned = LoadBitmap(hHandle, "#" & i)
    If lReturned <> 0 Then
        Form1.TreeView1.Nodes.Add "bmp", tvwChild, "B" & i, i & " [" & FileInfo(FileName).LanguageID & "]", 4, 4
    End If
Next i
FreeLibrary hHandle

End Function

Function LoadIconA(FileName As String, BitmapName As String, Pic As PictureBox)
Dim hInstance, lReturned, lRes, BMnewDC, newHDC, hHandle As Long
Dim lBuffer As String * 255
Dim MySize As Size

'loadimage
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
'lRes = FindResource(hHandle, BitmapName, RT_BITMAP)
'MsgBox GetLastError
hObject = LoadResource(hHandle, lRes)
lReturned = LoadIcon(hHandle, BitmapName)
'MsgBox "LoadBitmap returned " & lReturned
'MsgBox GetBitmapDimensionEx(lReturned, MySize)

'Pic.Height = MySize.cx '* 150
'Pic.Width = MySize.cy '* 15
'Debug.Print "cY = " & MySize.cy
'Debug.Print "cX = " & MySize.cx
Pic.Cls
'Pic.Visible = False
x = SelectObject(Pic.hdc, lReturned)
newdc = CreateCompatibleDC(Pic.hdc)
y = SelectObject(newdc, lReturned)
'Pic.Visible = True

FreeLibrary hHandle


'hHandle = LoadBitmap(hInstance, "#159")
End Function


Function GetIconsA(FileName As String)
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
For i = 1 To 9999
    lReturned = LoadIcon(hHandle, "#" & i)
    'If lReturned <> 0 Then Form1.List2.AddItem "#" & i
Next i
FreeLibrary hHandle

End Function





Function LoadCursorA(FileName As String, BitmapName As String, Pic As PictureBox)
Dim hInstance, lReturned, lRes, BMnewDC, newHDC, hHandle As Long
Dim lBuffer As String * 255
Dim MySize As Size

'loadimage
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
'lRes = FindResource(hHandle, BitmapName, RT_BITMAP)
'MsgBox GetLastError
hObject = LoadResource(hHandle, lRes)
lReturned = LoadCursor(hHandle, BitmapName)
'MsgBox "LoadBitmap returned " & lReturned
'MsgBox GetBitmapDimensionEx(lReturned, MySize)

'Pic.Height = MySize.cx '* 150
'Pic.Width = MySize.cy '* 15
'Debug.Print "cY = " & MySize.cy
'Debug.Print "cX = " & MySize.cx
Pic.Cls
'Pic.Visible = False
x = SelectObject(Pic.hdc, lReturned)
newdc = CreateCompatibleDC(Pic.hdc)
y = SelectObject(newdc, lReturned)

DrawIcon Pic.hdc, 1, 1, lReturned
'Pic.Visible = True
FreeLibrary hHandle
'MsgBox "SelectObject(Pic.hdc, lReturned) = " & newdc

'hHandle = LoadBitmap(hInstance, "#159")
End Function


Function GetCursorsA(FileName As String)
hHandle = LoadLibrary(FileName) 'GetModuleHandle(filename)
For i = 1 To 9999
    lReturned = LoadCursor(hHandle, "#" & i)
    If lReturned <> 0 Then
        Form1.TreeView1.Nodes.Add "cur", tvwChild, "C" & i, i & " [" & FileInfo(FileName).LanguageID & "]", 6, 6
    End If
Next i
FreeLibrary hHandle
End Function


Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
 ' return file-properties of given file  (EXE , DLL , OCX)

Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename
If Len(PathWithFilename) = 0 Then
    FileInfo = BACKUP
    Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
               lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         'Do While Len(strLangCharset) < 8
         '    strLangCharset = "0" & strLangCharset
         'Loop
         If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
         strLangCharset2 = "English (US)"
         'If Left(strLangCharset, 2) = SUBLANG_ENGLISH_US Then strLangCharset = "English (US)"
         'If Left(strLangCharset, 2) = SUBLANG_ENGLISH_CAN Then strLangCharset = "English (CANADIAN)"
         'If Left(strLangCharset, 2) = SUBLANG_ENGLISH_AUS Then strLangCharset = "English (AUSTRALIAN)"
         'If Left(strLangCharset, 2) = SUBLANG_ENGLISH_NZ Then strLangCharset = "English (NEW ZEALAND)"
         'If Left(strLangCharset, 2) = SUBLANG_ENGLISH_UK Then strLangCharset = "English (UK)"
         
         End If
         If Mid(strLangCharset, 2, 2) = LANG_BULGARIAN Then strLangCharset2 = "Bulgarian"
         If Mid(strLangCharset, 2, 2) = LANG_FRENCH Then strLangCharset2 = "French"
         If Mid(strLangCharset, 2, 2) = LANG_NEUTRAL Then strLangCharset2 = "Neutral"
         
         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         ' loop and get fileproperties
         For intTemp = 0 To 7
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = "?"
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type
FileInfo.CompanyName = strVersionInfo(0)
FileInfo.FileDescription = strVersionInfo(1)
FileInfo.FileVersion = strVersionInfo(2)
FileInfo.InternalName = strVersionInfo(3)
FileInfo.LegalCopyright = strVersionInfo(4)
FileInfo.OrigionalFileName = strVersionInfo(5)
FileInfo.ProductName = strVersionInfo(6)
FileInfo.ProductVersion = strVersionInfo(7)
FileInfo.LanguageID = strLangCharset2
BACKUP = FileInfo
End Function



