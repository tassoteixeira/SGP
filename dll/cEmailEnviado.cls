VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmailEnviado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumero As Long
Private mRemetente As String
Private mDestinatario As String
Private mEmailRemetente As String
Private mEmailDestinatario As String
Private mAssunto As String
Private mMensagem As String
Private mCriadoEm As String
Private mEnviadoEm As String
Private mIPRemetente As String
Private mEmailEnviado As Boolean
Private mEmailConcluido As Boolean

Private rsEmailEnviado As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
End Property


Public Property Get Remetente() As String
    Remetente = mRemetente
End Property
Public Property Let Remetente(ByVal Valor As String)
    mRemetente = Valor
End Property


Public Property Get Destinatario() As String
    Destinatario = mDestinatario
End Property
Public Property Let Destinatario(ByVal Valor As String)
    mDestinatario = Valor
End Property


Public Property Get EmailRemetente() As String
    EmailRemetente = mEmailRemetente
End Property
Public Property Let EmailRemetente(ByVal Valor As String)
    mEmailRemetente = Valor
End Property


Public Property Get EmailDestinatario() As String
    EmailDestinatario = mEmailDestinatario
End Property
Public Property Let EmailDestinatario(ByVal Valor As String)
    mEmailDestinatario = Valor
End Property


Public Property Get Assunto() As String
    Assunto = mAssunto
End Property
Public Property Let Assunto(ByVal Valor As String)
    mAssunto = Valor
End Property


Public Property Get Mensagem() As String
    Mensagem = mMensagem
End Property
Public Property Let Mensagem(ByVal Valor As String)
    mMensagem = Valor
End Property


Public Property Get CriadoEm() As String
    CriadoEm = mCriadoEm
End Property
Public Property Let CriadoEm(ByVal Valor As String)
    mCriadoEm = Valor
End Property


Public Property Get EnviadoEm() As String
    EnviadoEm = mEnviadoEm
End Property
Public Property Let EnviadoEm(ByVal Valor As String)
    mEnviadoEm = Valor
End Property


Public Property Get IPRemetente() As String
    IPRemetente = mIPRemetente
End Property
Public Property Let IPRemetente(ByVal Valor As String)
    mIPRemetente = Valor
End Property


Public Property Get EmailEnviado() As Boolean
    EmailEnviado = mEmailEnviado
End Property
Public Property Let EmailEnviado(ByVal Valor As Boolean)
    mEmailEnviado = Valor
End Property


Public Property Get EmailConcluido() As Boolean
    EmailConcluido = mEmailConcluido
End Property
Public Property Let EmailConcluido(ByVal Valor As Boolean)
    mEmailConcluido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pNumero As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Numero = " & pNumero
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Numero < " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Numero DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Numero ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Numero > " & mNumero
    gSQL = PreparaSQL(xCondicao, "ORDER BY Numero ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Numero DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mNumero = ProximoRegistro(mEmpresa)
    gSQL = "INSERT INTO EmailEnviado ( Empresa, Numero, Remetente, Destinatario, [Email Remetente], "
    gSQL = gSQL & "[Email Destinatario], Assunto, Mensagem, [Criado Em], [Enviado Em], "
    gSQL = gSQL & "[IP Remetente], [Email Enviado], [Email Concluido] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlTexto(1, mRemetente, ", ")
    Call sqlTexto(1, mDestinatario, ", ")
    Call sqlTexto(1, mEmailRemetente, ", ")
    Call sqlTexto(1, mEmailDestinatario, ", ")
    Call sqlTexto(1, mAssunto, ", ")
    Call sqlTexto(1, mMensagem, ", ")
    Call sqlTexto(1, mCriadoEm, ", ")
    Call sqlTexto(1, mEnviadoEm, ", ")
    Call sqlTexto(1, mIPRemetente, ", ")
    Call sqlBoolean(1, mEmailEnviado, ", ")
    Call sqlBoolean(1, mEmailConcluido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE EmailEnviado SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Numero = ", mNumero)
    Call sqlTexto(2, ", Remetente = ", mRemetente)
    Call sqlTexto(2, ", Destinatario = ", mDestinatario)
    Call sqlTexto(2, ", [Email Remetente] = ", mEmailRemetente)
    Call sqlTexto(2, ", [Email Destinatario] = ", mEmailDestinatario)
    Call sqlTexto(2, ", Assunto = ", mAssunto)
    Call sqlTexto(2, ", Mensagem = ", mMensagem)
    Call sqlTexto(2, ", [Criado Em] = ", mCriadoEm)
    Call sqlTexto(2, ", [Enviado Em] = ", mEnviadoEm)
    Call sqlTexto(2, ", [IP Remetente] = ", mIPRemetente)
    Call sqlBoolean(2, ", [Email Enviado] = ", mEmailEnviado)
    Call sqlBoolean(2, ", [Email Concluido] = ", mEmailConcluido)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pNumero As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM EmailEnviado"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Numero = " & pNumero
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Numero, Remetente, Destinatario, [Email Remetente], [Email Destinatario], Assunto, Mensagem, [Criado Em], [Enviado Em], [IP Remetente], [Email Enviado], [Email Concluido] FROM EmailEnviado" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEmailEnviado = New adodb.Recordset
    rsEmailEnviado.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEmailEnviado.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEmailEnviado.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEmailEnviado("Empresa").Value
    Numero = rsEmailEnviado("Numero").Value
    Remetente = rsEmailEnviado("Remetente").Value
    Destinatario = rsEmailEnviado("Destinatario").Value
    EmailRemetente = rsEmailEnviado("Email Remetente").Value
    EmailDestinatario = rsEmailEnviado("Email Destinatario").Value
    Assunto = rsEmailEnviado("Assunto").Value
    Mensagem = rsEmailEnviado("Mensagem").Value
    CriadoEm = rsEmailEnviado("Criado Em").Value
    EnviadoEm = rsEmailEnviado("Enviado Em").Value
    IPRemetente = rsEmailEnviado("IP Remetente").Value
    EmailEnviado = rsEmailEnviado("Email Enviado").Value
    EmailConcluido = rsEmailEnviado("Email Concluido").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function



Public Function ProximoRegistro(ByVal pEmpresa As Integer) As Long

On Error GoTo Err_ProximoRegistro

    'Pega número do último movimento
    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT Numero"
    gSQL = gSQL & "  FROM EmailEnviado"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Numero DESC"
    If Localizar(1) Then
        ProximoRegistro = rsEmailEnviado("Numero").Value + 1
    End If
    rsEmailEnviado.Close
    Set rsEmailEnviado = Nothing
    Exit Function

Err_ProximoRegistro:
End Function

