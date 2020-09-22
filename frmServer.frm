VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmServer 
   Caption         =   "WebServer"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frDomain 
      Caption         =   "[ Domain ]"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   7935
      Begin VB.CommandButton cmdDomainDelete 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDomainAdd 
         Caption         =   "Add new"
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwDomain 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Domain name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Root directory"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Default document"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame frServer 
      Caption         =   "[ Private Web Server ]"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtRoot 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         TabIndex        =   5
         Text            =   "F:\Joy\Web Projects\"
         Top             =   300
         Width           =   6135
      End
      Begin VB.CommandButton cmdServer 
         Caption         =   "Start"
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         TabIndex        =   3
         Text            =   "80"
         Top             =   1222
         Width           =   855
      End
      Begin VB.TextBox txtDefaultDocument 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1680
         TabIndex        =   2
         Text            =   "index.html;index.htm;default.html;default.htm"
         Top             =   780
         Width           =   6135
      End
      Begin MSWinsockLib.Winsock sckServer 
         Index           =   0
         Left            =   1200
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblRoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Root directory"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label lblDefaultDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default document"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1485
      End
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7435
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmServer.frx":0ECA
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDomainAdd_Click()
frmDomain.Show
End Sub

Private Sub cmdDomainDelete_Click()
If MsgBox("Are you sure you want to remove the selected domains? This action cannot be reversed.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

Dim lItem As ListItem
For Each lItem In lvwDomain.ListItems
    If lItem.Checked Then
        QueryExec "DELETE FROM tblDomain WHERE DomainID = " & Mid(lItem.Key, 8)
    End If
Next

LoadDomainList
End Sub

Private Sub cmdServer_Click()
Select Case cmdServer.Caption
Case "Start"
    Status "Starting server..."

    sckServer(0).LocalPort = Val(txtPort)
    sckServer(0).Listen
    
    If sckServer(0).State = sckListening Then
        txtPort.Enabled = False
        cmdServer.Caption = "Stop"
        Status "Server started"
    Else
        Status "Server could not be started"
        MsgBox "Sorry, the server could not be started. Current socket state is " & sckServer(0).State
    End If
Case "Stop"
    Status "Stopping server..."
    
    CloseSocket sckServer(0)
    
    If sckServer(0).State = sckClosed Then
        txtPort.Enabled = True
        cmdServer.Caption = "Start"
        Status "Server stopped"
    Else
        Status "Server could not be stopped"
        MsgBox "Sorry, the server could not be stopped. Current socket state is " & sckServer(0).State
    End If
    
End Select
End Sub

Private Sub Form_Load()
Me.Show
Me.Refresh

LoadDomainList

Status "Server ready..."
End Sub

Private Sub Form_Resize()
On Error Resume Next

frDomain.Move 8, frDomain.Top, Me.ScaleWidth - 16, Me.ScaleHeight - frDomain.Top - 8
lvwDomain.Move 120, 600, frDomain.Width * Screen.TwipsPerPixelX - 240, frDomain.Height * Screen.TwipsPerPixelY - 720
cmdDomainAdd.Move frDomain.Width * Screen.TwipsPerPixelX - cmdDomainAdd.Width - cmdDomainDelete.Width - 120
cmdDomainDelete.Move frDomain.Width * Screen.TwipsPerPixelX - cmdDomainDelete.Width - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdServer.Caption = "Stop" Then cmdServer_Click

UnloadAllForms
End Sub

Private Sub lvwDomain_DblClick()
Dim fDomain As frmDomain
Set fDomain = New frmDomain

fDomain.LoadDomain Val(Mid(lvwDomain.SelectedItem.Key, 8))
End Sub

Private Sub sckServer_Close(Index As Integer)
Status "Channel " & Index & " closed"
End Sub

Private Sub sckServer_Connect(Index As Integer)
Status "Browser connected at channel " & Index
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index > 0 Then Exit Sub

Dim CurrentSocketIndex As Long
CurrentSocketIndex = GetFreeSocketIndex
Load sckServer(CurrentSocketIndex)
sckServer(CurrentSocketIndex).Accept requestID

Status "Channel(" & CurrentSocketIndex & ") created for browser request..."
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Status "Data request at channel " & Index & "..."

Dim Request As String
sckServer(Index).GetData Request

Dim URL As String, DomainName As String, RootPath As String, Document As String, DefaultDocument As String

'Determine what kind of request it is
If UCase(Left(Request, 4)) = "GET " Then 'HTTP GET request
    While InStr(Request, "  ") > 0 'Rip off all the multiple spaces
        Request = Left(Request, InStr(Request, "  ") - 1) & Mid(Request, InStr(Request, "  ") + 1)
        DoEvents
    Wend
    
    Request = Mid(Request, InStr(Request, " ") + 1) 'Skip the command bytes
    Request = Left(Request, InStr(Request, " ") - 1)
    If UCase(Left(Request, 7)) = "HTTP://" Then Request = Mid(Request, 8)
    If Left(Request, 1) = "/" Then Request = Mid(Request, 2)
    If Left(Request, 1) = "/" Then Request = Mid(Request, 2)
    
    URL = Request 'Get the URL
    
    'Get the domain name
    If InStr(URL, "/") = 0 Then
        DomainName = URL
        URL = URL & "/"
    Else
        DomainName = Left(URL, InStr(URL, "/") - 1)
    End If
    
    'Get the template to serve
    Document = Mid(URL, InStr(URL, "/") + 1)
    Document = Replace(Document, "/", "\")
    Document = Replace(Document, "%20", " ")
        
    'Check if this server hosts the domain
    If Query("SELECT * FROM tblDomain WHERE DomainName = 'www." & DomainName & "' OR DomainName = '" & DomainName & "'").RecordCount > 0 Then
    'Yes, this server hosts the domain
    
        'So get the document root for the domain
        RootPath = Query("SELECT * FROM tblDomain WHERE DomainName = 'www." & DomainName & "' OR DomainName = '" & DomainName & "'").Fields!RootPath
        DefaultDocument = Query("SELECT * FROM tblDomain WHERE DomainName = 'www." & DomainName & "' OR DomainName = '" & DomainName & "'").Fields!DefaultDocument
        Document = RootPath & Document 'Full qualified path of the document
        
        Dim DDocCount As Long
        
        'Check if the document points to a file or a folder
        If Right(Document, 1) = "\" Then 'Serve the defualt document
            For DDocCount = 1 To ListLen(DefaultDocument, ";")
                If Dir(Document & ListGetAt(DefaultDocument, DDocCount, ";")) = "" Then 'The document not found
                    
                Else
                    'The defualt document found, so send it
                    SendDoc Document & ListGetAt(DefaultDocument, DDocCount, ";"), sckServer(Index)
                    Exit For 'No need to send any other default document
                End If
            Next
            'Show the error messege as no default document could be found to show
            If DDocCount = ListLen(DefaultDocument, ";") And Dir(Document & ListGetAt(DefaultDocument, DDocCount, ";")) = "" Then SendDoc CheckPath(App.Path, True) & "ErrorDoc\401.htm", sckServer(Index)
        Else
            If Dir(Document) = "" Then 'This is certainly not a file, so check again if it's a folder!
                Document = Document & "\"
                
                For DDocCount = 1 To ListLen(DefaultDocument, ";")
                    If Dir(Document & ListGetAt(DefaultDocument, DDocCount, ";")) = "" Then 'The document not found
                        
                    Else
                        'The defualt document found, so send it
                        SendDoc Document & ListGetAt(DefaultDocument, DDocCount, ";"), sckServer(Index)
                        Exit For 'No need to send any other default document
                    End If
                Next
                'Show the error messege as no default document could be found to show
                If DDocCount - 1 = ListLen(DefaultDocument, ";") And Dir(Document & ListGetAt(DefaultDocument, DDocCount - 1, ";")) = "" Then SendDoc CheckPath(App.Path, True) & "ErrorDoc\401.htm", sckServer(Index)
            Else 'The requested file is found, serve it
                SendDoc Document, sckServer(Index)
            End If
        End If
    Else
    'Bad request, the domain is not hosted on this server!
        SendDoc CheckPath(App.Path, True) & "ErrorDoc\404.htm", sckServer(Index)
    End If
Else 'Unknown request

End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SocketError = True

Status "Socket error#" & Number & "- " & Description
End Sub

Private Sub sckServer_SendComplete(Index As Integer)
Status "Data sending complete at channel " & Index

CloseSocket sckServer(Index)
Unload sckServer(Index)
End Sub

Private Sub txtDefaultDocument_LostFocus()
If Trim(txtDefaultDocument) = "" Then txtDefaultDocument = "index.htm;index.html"

Status "Default document list updated"
End Sub

Private Sub txtPort_LostFocus()
Status "Server port updated"

End Sub

Private Sub txtRoot_LostFocus()
If Right(txtRoot.Text, 1) <> "\" Then txtRoot = txtRoot & "\"

Status "Root directory updated"
End Sub

Private Function GetFreeSocketIndex() As Long
Status "Searching for free channel..."
Dim s As Winsock, c As Long
For Each s In sckServer
    If s.Index > 0 Then
        c = c + 1
        If s.Index > c Then
            GetFreeSocketIndex = c
            Exit Function
        End If
    End If
Next
GetFreeSocketIndex = c + 1
End Function

Private Sub SendDoc(Document As String, Socket As Winsock, Optional Path As String = "./")
If FileLen(Document) = 0 Then Exit Sub

Dim DataBuffer As Long

DataBuffer = 5000
Dim FileData As String * 5000 'Must size this string to the DataBuffer value

Dim FileNumber As Long, PositionCount As Long

FileNumber = FreeFile
Open Document For Binary As FileNumber

SocketSend Socket, "HTTP/1.1" & vbCrLf
SocketSend Socket, "Server: Private Web Server" & vbCrLf
SocketSend Socket, "Path=" & Path & vbCrLf
SocketSend Socket, vbCrLf

If LOF(FileNumber) < DataBuffer Then
    Get #FileNumber, , FileData
    SocketSend Socket, Left(FileData, LOF(FileNumber))
Else
    For PositionCount = 1 To LOF(FileNumber) Step DataBuffer
        Get #FileNumber, , FileData
        SocketSend Socket, FileData
    Next
    If LOF(FileNumber) > PositionCount - 1 Then
        Get #FileNumber, , FileData
        SocketSend Socket, Left(FileData, LOF(FileNumber) - (PositionCount - 1))
    End If
End If
Close #FileNumber
End Sub

Private Sub LoadDomainList()
lvwDomain.ListItems.Clear

With Query("SELECT * FROM tblDomain ORDER BY DomainName")
    While Not .EOF
        lvwDomain.ListItems.Add , "Domain#" & .Fields!DomainID, .Fields("DomainName")
        lvwDomain.ListItems("Domain#" & .Fields!DomainID).SubItems(1) = .Fields!RootPath
        lvwDomain.ListItems("Domain#" & .Fields!DomainID).SubItems(2) = .Fields!DefaultDocument
        
        .MoveNext
        DoEvents
    Wend
End With
End Sub
