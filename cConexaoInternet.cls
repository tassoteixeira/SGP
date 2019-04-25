VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConexaoInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long

Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40


Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
    Dim dwflags As Long
    Dim WebTest As Boolean
    ConnType = ""
    WebTest = InternetGetConnectedState(dwflags, 0&)
    Select Case WebTest
        Case dwflags And CONNECT_LAN: ConnType = "LAN"
        Case dwflags And CONNECT_MODEM: ConnType = "Modem"
        Case dwflags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwflags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwflags And CONNECT_CONFIGURED: ConnType = "Configurada"
        Case dwflags And CONNECT_RAS: ConnType = "Remota"
    End Select
IsWebConnected = WebTest
End Function

Private Sub Command1_Click()
    Dim msg As String
    If IsWebConnected(msg) Then
        msg = "Você está conectado à Internet via: " & msg
    Else
        msg = "Você não está conectado à Internet."
    End If
    MsgBox msg, vbOKOnly, "Estado de conexão à Internet"

End Sub

