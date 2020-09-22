Attribute VB_Name = "modWinSockFunction"
Option Explicit

Public SocketError As Boolean, RxBytes As Long, TxBytes As Long

Public Sub CloseSocket(sckObj As Winsock)
SocketError = False
sckObj.Close
While sckObj.State <> sckClosed And sckObj.State <> sckError And Not SocketError
    DoEvents
Wend
End Sub

Public Sub SocketSend(sckObj As Winsock, Data2Send As String)
If sckObj.State = sckConnected Then sckObj.SendData Data2Send

TxBytes = TxBytes + Len(Data2Send)
End Sub

Public Sub ConnectSocket(sckObj As Winsock, RemoteHost As String, RemotePort As String, Optional TimeOut As Long = 20)
Dim StartTime As Date

CloseSocket sckObj

sckObj.Connect RemoteHost, RemotePort
StartTime = Now
SocketError = False
Do While sckObj.State <> sckConnected And sckObj.State <> sckError And Not SocketError
    If DateDiff("s", StartTime, Now) > TimeOut Then
        Exit Do
    End If
    DoEvents
Loop
End Sub
