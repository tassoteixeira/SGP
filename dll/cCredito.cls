VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mTipoPessoa As Integer
Private mCodigo As Long
Private mAprovado As Boolean
Private mPrazo As Integer
Private mDiasAtraso As Integer
Private mLimite As Currency
Private mCondicao As String

Private rsCredito As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get TipoPessoa() As Integer
    TipoPessoa = mTipoPessoa
End Property
Public Property Let TipoPessoa(ByVal Valor As Integer)
    mTipoPessoa = Valor
End Property


Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Long)
    mCodigo = Valor
End Property


Public Property Get Aprovado() As Boolean
    Aprovado = mAprovado
End Property
Public Property Let Aprovado(ByVal Valor As Boolean)
    mAprovado = Valor
End Property


Public Property Get Prazo() As Integer
    Prazo = mPrazo
End Property
Public Property Let Prazo(ByVal Valor As Integer)
    mPrazo = Valor
End Property


Public Property Get DiasAtraso() As Integer
    DiasAtraso = mDiasAtraso
End Property
Public Property Let DiasAtraso(ByVal Valor As Integer)
    mDiasAtraso = Valor
End Property


Public Property Get Limite() As Currency
    Limite = mLimite
End Property
Public Property Let Limite(ByVal Valor As Currency)
    mLimite = Valor
End Property


Public Property Get Condicao() As String
    Condicao = mCondicao
End Property
Public Property Let Condicao(ByVal Valor As String)
    mCondicao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsCredito.Close
    Set rsCredito = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Credito ( [Tipo de Pessoa], Codigo, Aprovado, Prazo, [Dias de Atraso], "
    gSQL = gSQL & "Limite, Condicao ) VALUES ( "
    Call sqlNumero(1, mTipoPessoa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlBoolean(1, mAprovado, ", ")
    Call sqlNumero(1, mPrazo, ", ")
    Call sqlNumero(1, mDiasAtraso, ", ")
    Call sqlValor(1, mLimite, ", ")
    Call sqlTexto(1, mCondicao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigo As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Credito SET "
    Call sqlNumero(2, "[Tipo de Pessoa] = ", mTipoPessoa)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlBoolean(2, ", Aprovado = ", mAprovado)
    Call sqlNumero(2, ", Prazo = ", mPrazo)
    Call sqlNumero(2, ", [Dias de Atraso] = ", mDiasAtraso)
    Call sqlValor(2, ", Limite = ", mLimite)
    Call sqlTexto(2, ", Condicao = ", mCondicao)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigo As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Credito"
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Tipo de Pessoa], Codigo, Aprovado, Prazo, [Dias de Atraso], Limite, Condicao FROM Credito" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsCredito = New adodb.Recordset
    rsCredito.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCredito.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCredito.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    TipoPessoa = rsCredito("Tipo de Pessoa").Value
    Codigo = rsCredito("Codigo").Value
    Aprovado = rsCredito("Aprovado").Value
    Prazo = rsCredito("Prazo").Value
    DiasAtraso = rsCredito("Dias de Atraso").Value
    Limite = rsCredito("Limite").Value
    Condicao = rsCredito("Condicao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

