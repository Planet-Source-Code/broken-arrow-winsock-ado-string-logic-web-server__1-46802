Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Private Declare Function GetPrivateProfilestring& Lib "Kernel32" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String)
'Private Declare Function WritePrivateProfileString& Lib "Kernel32" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String)

Public Function CheckPath(PathString As String, Optional AddSlash As Boolean = True) As String
If PathString = "" Then
    CheckPath = CheckPath(App.Path, True)
    Exit Function
End If

If AddSlash And Right(PathString, 1) <> "\" Then CheckPath = PathString & "\"
If Not AddSlash And Right(PathString, 1) = "\" Then CheckPath = Left(PathString, Len(PathString) - 1)
End Function

Public Sub GrowForm(frmObj As Form, Optional FramesPerSec As Long = 600)
Dim vWidth As Long, vHeight As Long, Count As Long

Load frmObj

vWidth = frmObj.Width
vHeight = frmObj.Height

With frmObj
    frmObj.Move -1 * frmObj.Width, -1 * frmObj.Height
    frmObj.Show
    
    For Count = 1 To vWidth Step FramesPerSec
        frmObj.Move (Screen.Width - Count) / 2, (Screen.Height - (vHeight * Count / vWidth)) / 2, Count, vHeight * Count / vWidth
        frmObj.Refresh
        DoEvents
    Next
    
    frmObj.Move (Screen.Width - vWidth) / 2, (Screen.Height - vHeight) / 2, vWidth, vHeight
End With
End Sub

Public Sub ShrinkForm(frmObj As Form, Optional FramesPerSec As Long = 600, Optional UnloadForm As Boolean = False)
Dim vWidth As Long, vHeight As Long, Count As Long

vWidth = frmObj.Width
vHeight = frmObj.Height

With frmObj
    For Count = vWidth To 1 Step -1 * FramesPerSec
        frmObj.Move (Screen.Width - Count) / 2, (Screen.Height - (vHeight * Count / vWidth)) / 2, Count, vHeight * Count / vWidth
        frmObj.Refresh
        DoEvents
    Next
    
    frmObj.Hide
    If UnloadForm Then Unload frmObj
End With
End Sub

Public Sub UnloadAllForms()
Dim frmObj As Form

For Each frmObj In Forms
    Unload frmObj
Next
End Sub

Public Function Encrypt(StringToEncrypt As String) As String
Dim a As Long, TempString1 As String, TempString2 As String
'Convert each byte to its decimal ASCII code, preserve 3 digit format with leading 0
For a = 1 To Len(StringToEncrypt)
    TempString1 = String(3 - Len(CStr(Asc(Mid(StringToEncrypt, a, 1)))), "0") & CStr(Asc(Mid(StringToEncrypt, a, 1))) & TempString1
Next

'Convert each byte to its hexadecimal ASCII code, preserve 2 digit format with leading 0
For a = 1 To Len(TempString1)
    TempString2 = TempString2 & String(2 - Len(Hex(Asc(Mid(TempString1, a, 1)))), "0") & Hex(Asc(Mid(TempString1, a, 1)))
Next

'Shift all odd placeholder digits to the left and all even placeholder digits to the right
TempString1 = TempString2
For a = 1 To Len(TempString2) / 2
    Mid(TempString1, a, 1) = Mid(TempString2, a * 2 - 1, 1)
    Mid(TempString1, a + Len(TempString2) / 2, 1) = Mid(TempString2, a * 2, 1)
Next

'Take each 2 digits from left and replace them with the ASCII character generated by them together
TempString2 = ""
For a = 1 To Len(TempString1) / 2
    TempString2 = TempString2 & Chr(Val(Mid(TempString1, a * 2 - 1, 2)))
Next

Encrypt = TempString2
End Function

Public Function Decrypt(StringToDecrypt As String) As String
Dim a As Long, TempString1 As String, TempString2 As String

For a = 1 To Len(StringToDecrypt)
    TempString1 = TempString1 & String(2 - Len(CStr(Asc(Mid(StringToDecrypt, a, 1)))), "0") & CStr(Asc(Mid(StringToDecrypt, a, 1)))
Next

For a = 1 To Len(TempString1) / 2
    TempString2 = TempString2 & Mid(TempString1, a, 1) & Mid(TempString1, Len(TempString1) / 2 + a, 1)
Next

TempString1 = ""
For a = 1 To Len(TempString2) / 2
    TempString1 = TempString1 & Chr(Hex2Dec(Mid(TempString2, a * 2 - 1, 2)))
Next

TempString2 = ""
For a = 1 To Len(TempString1) / 3
    TempString2 = Chr(Val(Mid(TempString1, a * 3 - 2, 3))) & TempString2
Next
End Function

Public Function Hex2Dec(HexValue As String) As Double
Dim a As Long, b As Double
HexValue = UCase(HexValue)
For a = 1 To Len(HexValue) - 1
    If Mid(HexValue, a, 1) >= "0" And Mid(HexValue, a, 1) <= "9" Then b = b + Val(Mid(HexValue, a, 1)) * 16 ^ ((Len(HexValue) - a + 1) - 1) Else b = b + (Asc(Mid(HexValue, a, 1)) - 55) * 16 ^ ((Len(HexValue) - a + 1) - 1)
Next
If Right(HexValue, 1) >= "0" And Right(HexValue, 1) <= "9" Then b = b + Val(Right(HexValue, 1)) Else b = b + Asc(Right(HexValue, 1)) - 55
Hex2Dec = b
End Function

Public Function GetFileFromFullPath(Full_Path As String) As String 'Extract only filename from full
Dim AltStr1 As String, AltStr2 As String, a As Long         'path & filename
For a = Len(Full_Path) To 1 Step -1
    AltStr1 = AltStr1 & Mid(Full_Path, a, 1)
Next
AltStr1 = Left(AltStr1, InStr(AltStr1, "\") - 1)
For a = Len(AltStr1) To 1 Step -1
    AltStr2 = AltStr2 & Mid(AltStr1, a, 1)
Next
GetFileFromFullPath = AltStr2
End Function

Public Function GetPathFromFullPath(Full_Path As String) As String 'Extract path from a full path
GetPathFromFullPath = Left(Full_Path, Len(Full_Path) - Len(GetFileFromFullPath(Full_Path)) - 1)
End Function

Public Function IsItemInList(Item As String, List_Object As Object) As Boolean 'Check if an item is
Dim a As Long, Found As Boolean                                         'already in a list
If List_Object.ListCount = 0 Then Exit Function                         'control's list
For a = 0 To List_Object.ListCount - 1
    If List_Object.List(a) = Item Then
        Found = True
        Exit For
    End If
Next
IsItemInList = Found
End Function

Public Sub CutCopyPaste(ActionChoice As Integer)
'ActiveForm refers to the active form in the MDI form.
If TypeOf Screen.ActiveControl Is TextBox Then
    Select Case ActionChoice
        Case 0 ' Cut.
            Clipboard.SetText Screen.ActiveControl.SelText
            Screen.ActiveControl.SelText = ""
        Case 1 ' Copy.
            Clipboard.SetText Screen.ActiveControl.SelText
        Case 2 ' Paste.
            Screen.ActiveControl.SelText = Clipboard.GetText()
        Case 3 ' Delete.
            Screen.ActiveControl.SelText = ""
    End Select
End If
End Sub

Public Sub MakeDir(DirectoryName As String)
Dim iMouseState As Integer
Dim iNewLen As Integer
Dim iDirLen As Integer

iMouseState = Screen.MousePointer 'Get Mouse State
Screen.MousePointer = 11 'Change Mouse To Hour Glass
iNewLen = 4 'Set Start Length To Search For [\]
If Right$(DirectoryName, 1) <> "\" Then DirectoryName = DirectoryName + "\" 'Add [\] To Directory Name If Not There
While Not ValDir(DirectoryName) 'Create Nested Directory
    iDirLen = InStr(iNewLen, DirectoryName, "\")
    If Not ValDir(Left$(DirectoryName, iDirLen)) Then MkDir Left$(DirectoryName, iDirLen - 1)
    iNewLen = iDirLen + 1
    DoEvents
Wend
Screen.MousePointer = iMouseState 'Leave The Mouse The Way You Found It
End Sub

Private Function ValDir(sIncoming As String) As Integer
On Local Error GoTo ValDirError

Dim iCheck As String, iErrResult As Integer

iCheck = Dir$(sIncoming)
If iErrResult = 76 Then ValDir = False Else ValDir = True
Exit Function

ValDirError:
    Select Case Err
        Case Is = 76
           iErrResult = Err
           Resume Next
        Case Else
    End Select
End Function

Public Function ReadINIStr(SectionName As String, KeyName As String, Optional FileName As String = "", Optional DefaultValue As String) As String
Dim sRet As String
sRet = String(255, Chr(0))
If FileName = "" Then FileName = CheckPath(App.Path, True) & App.EXEName & ".INI"
GetPrivateProfileString SectionName, ByVal KeyName, DefaultValue, sRet, Len(sRet), FileName
sRet = Replace(sRet, Chr(0), "")
ReadINIStr = Trim(sRet)
End Function

Public Function ReadINIInt(SectionName As String, KeyName As String, Optional FileName As String = "", Optional DefaultValue As Long) As Double
If FileName = "" Then FileName = CheckPath(App.Path, True) & App.EXEName & ".INI"
ReadINIInt = GetPrivateProfileInt(SectionName, KeyName, DefaultValue, FileName)
End Function

Public Sub WriteINI(SectionName As String, KeyName As String, NewString As String, Optional FileName As String = "")
If FileName = "" Then FileName = CheckPath(App.Path, True) & App.EXEName & ".INI"
WritePrivateProfileString SectionName, KeyName, NewString, FileName
End Sub

Public Sub SelectAllText(ctl As Control, Optional ctlSetFocus As Boolean = False)
With ctl
    .SelStart = 0
    .SelLength = Len(ctl.Text)
    If ctlSetFocus Then .SetFocus
End With
End Sub

Public Function Contains(String1 As String, String2 As String) As Long
Contains = 0

Dim Count As Long
For Count = 1 To Len(String2)
    If InStr(String1, Mid(String2, Count, 1)) > 0 Then
        Contains = Count
        Exit For
    End If
Next
End Function

Public Function ContainsOtherThan(String1 As String, String2 As String) As Long
Dim a As Long

For a = 1 To Len(String1)
    If InStr(String2, Mid(String1, a, 1)) = 0 Then
        ContainsOtherThan = a
        Exit For
    End If
Next
End Function