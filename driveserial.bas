Attribute VB_Name = "Drive_Serial"
Option Explicit
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public Const MAX_FILENAME_LEN = 256
Public Function DriveSerial(ByVal sDrv As String) As Long
    Dim retval As Long
    Dim str As String * MAX_FILENAME_LEN
    Dim str2 As String * MAX_FILENAME_LEN
    Dim a As Long
    Dim b As Long
    Call GetVolumeInformation(sDrv & ":\", str, MAX_FILENAME_LEN, retval, a, b, str2, MAX_FILENAME_LEN)
    DriveSerial = retval
End Function


