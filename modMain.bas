Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
SetDatabase CheckPath(App.Path, True) & "PWS.mdb"

Load frmServer

End Sub

Public Sub Status(stsText As String)
With frmServer.txtStatus
    .SelStart = Len(.Text) + 1
    .SelText = stsText & vbCrLf
End With
End Sub
