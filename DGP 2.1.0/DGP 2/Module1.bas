Attribute VB_Name = "Module1"
Public corrfunc As String
Public corrtypefunc As Boolean
Public histstack As Byte
Public colorsrnd(7) As Long
Public clrnum As Byte
Public X, Y As Double
Public tracer As Boolean
Public allcount As Single
Public Const WM_PAINT = &HF
Public Const WM_PRINT = &H317
Public Const PRF_CLIENT = &H4&    ' Draw the window's client area.
Public Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
Public Const PRF_OWNED = &H20&    ' Draw all owned windows.
Public xmin, xmax, ymin As Single ' These variables used for main form
Public ymax, tempval As Single    ' and for detailed form
Public tracehist(7) As String
Public traceequation(7) As String 'These variable must be passed from main
Public tracex(7) As Boolean       'form to detailed form
Public traceeq As Boolean         'this variable should be passed to trace form
Public equationhistory(7) As String 'These variables used for main form
Public eqhistorytype(7) As Boolean 'and for detailed form
Public drawline As Boolean
Public langval As Byte                  'current language
Public Const pi = 3.14159265358979 ' declare value of pi
'use API for printing declare sendmessage
Public Declare Function SendMessage Lib "user32" Alias _
      "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long) As Long
' use API to read from file
Public Declare Function GetPrivateProfileString Lib _
     "kernel32" Alias "GetPrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal _
     lpKeyName As Any, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, ByVal lpFileName As String) _
     As Long
' use API to write into file
Public Declare Function WritePrivateProfileString Lib _
     "kernel32" Alias "WritePrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal _
     lpKeyName As Any, ByVal lpString As Any, ByVal _
     lpFileName As String) As Long
' sub write info for writing into .ini file
Public Sub WriteIniInfo(iniSection As String, _
                        iniItem As String, ItemValue As String, iniFile)
    Dim X As Long
    X = WritePrivateProfileString(iniSection, iniItem, ItemValue, iniFile)
    Exit Sub
End Sub
' function to read from file and update settings
Public Function GetIniInfo(iniFile As String, Section _
                           As String, ItemReturn As String, _
                           DefaultValue As String) As String
    Dim lResult As Long
    Dim sIniString As String
    sIniString = String(20, 0)
    lResult = GetPrivateProfileString(Section, ItemReturn, _
                        DefaultValue, sIniString, Len(sIniString), iniFile)
    sIniString = Left$(sIniString, InStr(sIniString, Chr$(0)) - 1)
    GetIniInfo = sIniString
    Exit Function
End Function

Public Function clr()
colorsrnd(0) = &H80&
colorsrnd(1) = &H800000
colorsrnd(2) = &H8000&
colorsrnd(3) = &H808000
colorsrnd(4) = &H8080&
colorsrnd(5) = &H800080
colorsrnd(6) = &H40C0&
colorsrnd(7) = &H0
End Function

