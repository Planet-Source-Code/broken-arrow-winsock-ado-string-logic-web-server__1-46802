VERSION 5.00
Begin VB.Form frmDomain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Domain Management"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDomain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Add new"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDefaultDocument 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1800
      TabIndex        =   5
      Text            =   "index.html;index.htm;default.asp;index.asp;default.php;index.php;index.cfm"
      Top             =   1020
      Width           =   5775
   End
   Begin VB.TextBox txtRootPath 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   540
      Width           =   5775
   End
   Begin VB.TextBox txtDomainName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Text            =   "www."
      Top             =   60
      Width           =   2895
   End
   Begin VB.Label lblDefaultDocument 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defualt document"
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label lblRootPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Root directory"
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1230
   End
   Begin VB.Label lblDomainName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domain name"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DomainIDToUpdate As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

If Trim(txtRootPath) = "" Then
    MsgBox "Please provide valid inputs and try again."
    Exit Sub
End If

If cmdOk.Caption = "Update" Then 'Edit the domain
    If Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "' AND DomainID <> " & DomainIDToUpdate).RecordCount > 0 Then
        MsgBox "Sorry, the domain already exists!"
        Exit Sub
    End If
    
    QueryExec "UPDATE tblDomain SET DomainName = '" & txtDomainName & "', RootPath = '" & txtRootPath & "', DefaultDocument = '" & txtDefaultDocument & "' WHERE DomainID = " & DomainIDToUpdate
    frmServer.lvwDomain.ListItems("Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID).Text = txtDomainName
    frmServer.lvwDomain.ListItems("Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID).SubItems(1) = txtRootPath
    frmServer.lvwDomain.ListItems("Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID).SubItems(2) = txtDefaultDocument
Else 'Add the domain
    If Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").RecordCount > 0 Then
        MsgBox "Sorry, the domain already exists!"
        Exit Sub
    End If
    
    QueryExec "INSERT INTO tblDomain (DomainName, RootPath, DefaultDocument) VALUES ('" & txtDomainName & "', '" & txtRootPath & "', '" & txtDefaultDocument & "')"
    frmServer.lvwDomain.ListItems.Add , "Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID, txtDomainName
    frmServer.lvwDomain.ListItems("Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID).SubItems(1) = txtRootPath
    frmServer.lvwDomain.ListItems("Domain#" & Query("SELECT * FROM tblDomain WHERE DomainName = '" & txtDomainName & "'").Fields!DomainID).SubItems(2) = txtDefaultDocument
End If

Unload Me
End Sub

Public Sub LoadDomain(DomainID As Long)
DomainIDToUpdate = DomainID

With Query("SELECT * FROM tblDomain WHERE DomainID = " & DomainID)
    txtDomainName = .Fields!DomainName
    txtRootPath = .Fields!RootPath
    txtDefaultDocument = .Fields!DefaultDocument
End With

cmdOk.Caption = "Update"

Me.Show
End Sub

Private Sub txtRootPath_LostFocus()
If Right(txtRootPath, 1) <> "\" Then txtRootPath = txtRootPath & "\"
End Sub
