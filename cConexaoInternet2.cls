VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConexaoInternet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const ERROR_SUCCESS = 0&
Private Const APINULL = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private ReturnCode As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Function ActiveConnection() As Boolean
  Dim hKey As Long
  Dim lpSubKey As String
  Dim phkResult As Long
  Dim lpValueName As String
  Dim lpReserved As Long
  Dim lpType As Long
  Dim lpData As Long
  Dim lpcbData As Long
  ActiveConnection = False
  lpSubKey = "System\CurrentControlSet\" & "Services\RemoteAccess"

  ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)

  If ReturnCode = ERROR_SUCCESS Then
    hKey = phkResult
    lpValueName = "Remote Connection"
    lpReserved = APINULL
    lpType = APINULL
    lpData = APINULL
    lpcbData = APINULL

    ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)

    lpcbData = Len(lpData)

    ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)

        
    If ReturnCode = ERROR_SUCCESS Then
      If lpData = 0 Then
        ActiveConnection = False
      Else
        ActiveConnection = True
      End If
    End If
  RegCloseKey (hKey)
  End If
End Function



