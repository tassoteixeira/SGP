Attribute VB_Name = "Leitura_dr_10"
'Declara��o da DLL com suas Fun��es
Public Declare Function IniPortaStr Lib "dr1032.DLL" (ByVal Porta As String, ByVal Velocidade As Integer) As Integer
Public Declare Function DRCarrega Lib "dr1032.DLL" () As Integer
Public Declare Function FechaPorta Lib "dr1032.DLL" () As Integer
 'Vari�vel para Controlar retorno da DLL
Global Retorno As Integer
 'Vari�vel que ir� conter o Comando a ser enviado para a Impressora
Global Comando As String
'*************************************************************
'Objetivo  -> Abrir a Porta de comunica��o e Verificar o Retorno da DLL
'**************************************************************
Public Function abre_porta()
'    If Form1!COM1.Value = True Then Retorno = IniPortaStr(Form1!COM1.Name) Else Retorno = IniPortaStr(Form1!COM2.Name)
    Retorno = IniPortaStr("COM2", 9600)
    If Retorno <> 1 Then
        MsgBox "Problemas ao Abrir a Porta de Comunica��o"
        fechar_porta
    End If
 End Function
'*************************************************************
'Objetivo  -> Fechar a Porta de Comunica��o e Verificar o Retorno da DLL
'**************************************************************
Public Function fechar_porta()
    Retorno = FechaPorta()
    If Retorno <> 1 Then
        MsgBox "Problemas ao Fechar a Porta de Comunica��o"
    End If
End Function
