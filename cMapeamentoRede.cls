VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMapeamentoRede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lRetorno As Long
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Public Function CriaConexao(ByVal pEndereco As String, ByVal pSenha As String, ByVal pUnidade As String) As Boolean
'pCaminho = "\\192.168.1.11\C"
'pUnidade = "Z:"
    CriaConexao = False
    lRetorno = WNetAddConnection(pEndereco, pSenha, pUnidade)
    If lRetorno = 0 Then
        CriaConexao = True
    End If
End Function
Public Function ConsultaConexao(ByVal pUnidade As String) As Boolean
    Dim xLocalAtual As String
'Consulta caminho de uma determinada conexão, _
'neste caso, na variável caminho será armazenado o local do caminho _
'atual ref. a unidade
    
    ConsultaConexao = False
    xLocalAtual = ""
    lRetorno = WNetGetConnection(pUnidade, xLocalAtual, 255)
    If lRetorno = 0 Then
        ConsultaConexao = True
    End If
End Function
Public Function Desconecta(ByVal pUnidade As String) As Boolean
    Desconecta = False
    lRetorno = WNetCancelConnection(pUnidade, True)
    If lRetorno = 0 Then
        Desconecta = True
    End If
End Function

