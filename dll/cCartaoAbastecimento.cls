VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCartaoAbastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroCartao As String
Private mCodigoFuncionario As Integer
Private mPeriodo As Integer
Private mPosicaoRegistro As String
Private mNumeroCartaoAPrazo As String
Private mPosicaoRegistroAPrazo As String


Private rsCartaoAbastecimento As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroCartao() As String
    NumeroCartao = mNumeroCartao
End Property
Public Property Let NumeroCartao(ByVal Valor As String)
    mNumeroCartao = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get PosicaoRegistro() As String
    PosicaoRegistro = mPosicaoRegistro
End Property
Public Property Let PosicaoRegistro(ByVal Valor As String)
    mPosicaoRegistro = Valor
End Property

Public Property Get NumeroCartaoAPrazo() As String
    NumeroCartaoAPrazo = mNumeroCartaoAPrazo
End Property
Public Property Let NumeroCartaoAPrazo(ByVal Valor As String)
    mNumeroCartaoAPrazo = Valor
End Property

Public Property Get PosicaoRegistroAPrazo() As String
    PosicaoRegistroAPrazo = mPosicaoRegistroAPrazo
End Property
Public Property Let PosicaoRegistroAPrazo(ByVal Valor As String)
    mPosicaoRegistroAPrazo = Valor
End Property


'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigoFuncionario(ByVal pEmpresa As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigoFuncionario = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigoFuncionario = Localizar(1)
    If LocalizarCodigoFuncionario = True Then
        AtribuiValor
    End If
    rsCartaoAbastecimento.Close
    Set rsCartaoAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroCartao(ByVal pEmpresa As Integer, ByVal pNumeroCartao As String, Optional ByVal pVerificaNumeroCartaoAPrazo As Boolean = False) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroCartao = False
    If pEmpresa > 0 Then
        xCondicao = " WHERE Empresa = " & pEmpresa
        xCondicao = xCondicao & " AND ([Numero do Cartao] = " & preparaTexto(pNumeroCartao)
    Else
        xCondicao = " WHERE ([Numero do Cartao] = " & preparaTexto(pNumeroCartao)
    End If
    
    If pVerificaNumeroCartaoAPrazo = True Then
        xCondicao = xCondicao & " OR [Numero do Cartao a Prazo] = " & preparaTexto(pNumeroCartao) & ")"
    Else
        xCondicao = xCondicao & ")"
    End If
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNumeroCartao = Localizar(1)
    If LocalizarNumeroCartao = True Then
        AtribuiValor
    End If
    rsCartaoAbastecimento.Close
    Set rsCartaoAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPosicaoRegistro(ByVal pEmpresa As Integer, ByVal pPosicaoRegistro As Integer, Optional ByVal pVerificaPosicaoAPrazo As Boolean = False) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPosicaoRegistro = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND ([Posicao do Registro] = " & preparaTexto(Format(pPosicaoRegistro, "000000"))
    
    If pVerificaPosicaoAPrazo = True Then
        xCondicao = xCondicao & " OR [Posicao do Registro a Prazo] = " & preparaTexto(Format(pPosicaoRegistro, "000000")) & ")"
    Else
        xCondicao = xCondicao & ")"
    End If
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarPosicaoRegistro = Localizar(1)
    If LocalizarPosicaoRegistro = True Then
        AtribuiValor
    End If
    rsCartaoAbastecimento.Close
    Set rsCartaoAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO CartaoAbastecimento ( Empresa, [Numero do Cartao], [Codigo do Funcionario], "
    gSQL = gSQL & "Periodo, [Posicao do Registro], [Numero do Cartao a Prazo], [Posicao do Registro a Prazo] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mNumeroCartao, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlTexto(1, mPosicaoRegistro, ", ")
    Call sqlTexto(1, mNumeroCartaoAPrazo, ", ")
    Call sqlTexto(1, mPosicaoRegistroAPrazo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pNumeroCartao As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE CartaoAbastecimento SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Numero do Cartao] = ", mNumeroCartao)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Posicao do Registro] = ", mPosicaoRegistro)
    Call sqlTexto(2, ", [Numero do Cartao a Prazo] = ", mNumeroCartaoAPrazo)
    Call sqlTexto(2, ", [Posicao do Registro a Prazo] = ", mPosicaoRegistroAPrazo)
    
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cartao] = " & preparaTexto(pNumeroCartao)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pNumeroCartao As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM CartaoAbastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cartao] = " & preparaTexto(pNumeroCartao)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Cartao], [Codigo do Funcionario], Periodo, [Posicao do Registro], [Numero do Cartao a Prazo], [Posicao do Registro a Prazo] FROM CartaoAbastecimento" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsCartaoAbastecimento = New adodb.Recordset
    rsCartaoAbastecimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCartaoAbastecimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCartaoAbastecimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsCartaoAbastecimento("Empresa").Value
    NumeroCartao = rsCartaoAbastecimento("Numero do Cartao").Value
    CodigoFuncionario = rsCartaoAbastecimento("Codigo do Funcionario").Value
    Periodo = rsCartaoAbastecimento("Periodo").Value
    If IsNull(rsCartaoAbastecimento("Posicao do Registro").Value) Then
        PosicaoRegistro = ""
    Else
        PosicaoRegistro = rsCartaoAbastecimento("Posicao do Registro").Value
    End If
    
    NumeroCartaoAPrazo = rsCartaoAbastecimento("Numero do Cartao a Prazo").Value
    
    If IsNull(rsCartaoAbastecimento("Posicao do Registro a Prazo").Value) Then
        PosicaoRegistroAPrazo = ""
    Else
        PosicaoRegistroAPrazo = rsCartaoAbastecimento("Posicao do Registro a Prazo").Value
    End If
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

