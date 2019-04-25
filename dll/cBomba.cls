VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBomba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As Integer
Private mTipoCombustivel As String
Private mPrecoCusto As Currency
Private mPrecoVenda As Currency
Private mNumeroTanque As String
Private mNumeroIlha As Integer
Private mTipoPreco As String
Private mCodigoProduto As Long
Private mNumeroSerie As String
Private mCodigoFisicoBomba As Integer
Private mCodigoFisicoBico As Integer
Private mLadoFisicoBico As Integer
Private mPontoAtendimentoBomba As Integer
Private mLadoFisicoPA As Integer
Private mPrecoVendaPrazo As Currency

Private rsBomba As New adodb.Recordset


'Inicio Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get NumeroTanque() As String
    NumeroTanque = mNumeroTanque
End Property
Public Property Let NumeroTanque(ByVal Valor As String)
    mNumeroTanque = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get TipoPreco() As String
    TipoPreco = mTipoPreco
End Property
Public Property Let TipoPreco(ByVal Valor As String)
    mTipoPreco = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get NumeroSerie() As String
    NumeroSerie = mNumeroSerie
End Property
Public Property Let NumeroSerie(ByVal Valor As String)
    mNumeroSerie = Valor
End Property


Public Property Get CodigoFisicoBomba() As Integer
    CodigoFisicoBomba = mCodigoFisicoBomba
End Property
Public Property Let CodigoFisicoBomba(ByVal Valor As Integer)
    mCodigoFisicoBomba = Valor
End Property


Public Property Get CodigoFisicoBico() As Integer
    CodigoFisicoBico = mCodigoFisicoBico
End Property
Public Property Let CodigoFisicoBico(ByVal Valor As Integer)
    mCodigoFisicoBico = Valor
End Property


Public Property Get LadoFisicoBico() As Integer
    LadoFisicoBico = mLadoFisicoBico
End Property
Public Property Let LadoFisicoBico(ByVal Valor As Integer)
    mLadoFisicoBico = Valor
End Property


Public Property Get PontoAtendimentoBomba() As Integer
    PontoAtendimentoBomba = mPontoAtendimentoBomba
End Property
Public Property Let PontoAtendimentoBomba(ByVal Valor As Integer)
    mPontoAtendimentoBomba = Valor
End Property

Public Property Get LadoFisicoPA() As Integer
    LadoFisicoPA = mLadoFisicoPA
End Property
Public Property Let LadoFisicoPA(ByVal Valor As Integer)
    mLadoFisicoPA = Valor
End Property

Public Property Get PrecoVendaPrazo() As Currency
    PrecoVendaPrazo = mPrecoVendaPrazo
End Property
Public Property Let PrecoVendaPrazo(ByVal Valor As Currency)
    mPrecoVendaPrazo = Valor
End Property


'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarCodigo = False
    xCondicao = "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCodigo = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigoProduto(ByVal nEmpresa As Integer, ByVal nCodigo As Long) As String
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarCodigoProduto = ""
    xCondicao = "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCodigoProduto = rsBomba("Tipo de Combustivel").Value
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarTipoCombustivel(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarTipoCombustivel = False
    xCondicao = "WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, " ORDER BY [Codigo do Produto]")
    If Localizar(1) Then
        LocalizarTipoCombustivel = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarTipoCombustivelPreco(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pTipoPreco As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarTipoCombustivelPreco = False
    xCondicao = "WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Tipo de Preco] = " & preparaTexto(pTipoPreco)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    If Localizar(1) Then
        LocalizarTipoCombustivelPreco = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarAnterior = False
    xCondicao = "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    If Localizar(1) Then
        LocalizarAnterior = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarPrimeiro = False
    xCondicao = "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    If Localizar(1) Then
        LocalizarPrimeiro = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarProximo = False
    xCondicao = "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    If Localizar(1) Then
        LocalizarProximo = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarUltimo = False
    xCondicao = "WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Codigo < 9999"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsBomba.Close
    Set rsBomba = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Bomba ( Empresa, Codigo, [Tipo de Combustivel], [Preco de Custo], "
    gSQL = gSQL & " [Preco de Venda], [Numero do Tanque], [Numero da Ilha], [Tipo de Preco], "
    gSQL = gSQL & " [Codigo do Produto], [Numero de Serie], [Codigo Fisico da Bomba],"
    gSQL = gSQL & " [Codigo Fisico do Bico], [Lado Fisico do Bico], [Ponto de Atendimento da Bomba], "
    gSQL = gSQL & " [Lado Fisico do PA], [Preco Venda Prazo] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlValor4(1, mPrecoCusto, ", ")
    Call sqlValor4(1, mPrecoVenda, ", ")
    Call sqlTexto(1, mNumeroTanque, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlTexto(1, mTipoPreco, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlTexto(1, mNumeroSerie, ", ")
    Call sqlNumero(1, mCodigoFisicoBomba, ", ")
    Call sqlNumero(1, mCodigoFisicoBico, ", ")
    Call sqlNumero(1, mLadoFisicoBico, ", ")
    Call sqlNumero(1, mPontoAtendimentoBomba, ", ")
    Call sqlNumero(1, mLadoFisicoPA, ", ")
    Call sqlValor4(1, mPrecoVendaPrazo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Bomba SET "
    'Call sqlNumero(2, "Empresa = ", mEmpresa)
    'Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, "[Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlValor4(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor4(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlTexto(2, ", [Numero do Tanque] = ", mNumeroTanque)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlTexto(2, ", [Tipo de Preco] = ", mTipoPreco)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlTexto(2, ", [Numero de Serie] = ", mNumeroSerie)
    Call sqlNumero(2, ", [Codigo Fisico da Bomba] = ", mCodigoFisicoBomba)
    Call sqlNumero(2, ", [Codigo Fisico do Bico] = ", mCodigoFisicoBico)
    Call sqlNumero(2, ", [Lado Fisico do Bico] = ", mLadoFisicoBico)
    Call sqlNumero(2, ", [Ponto de Atendimento da Bomba] = ", mPontoAtendimentoBomba)
    Call sqlNumero(2, ", [Lado Fisico do PA] = ", mLadoFisicoPA)
    Call sqlValor4(2, ", [Preco Venda Prazo] = ", mPrecoVendaPrazo)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarCodigoProduto(ByVal nEmpresa As Integer, ByVal nCodigoProduto As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarCodigoProduto

    AlterarCodigoProduto = False
    gSQL = "UPDATE Bomba SET "
    'Call sqlValor4(2, "[Preco de Custo] = ", mPrecoCusto)
    Call sqlValor4(2, ", [Preco de Venda] = ", mPrecoVenda)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Produto] = " & nCodigoProduto
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarCodigoProduto = True
    End If
    Exit Function

Err_AlterarCodigoProduto:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Bomba"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, [Tipo de Combustivel], [Preco de Custo], [Preco de Venda], [Numero do Tanque], [Numero da Ilha], [Tipo de Preco], [Codigo do Produto], [Numero de Serie], [Codigo Fisico da Bomba], [Codigo Fisico do Bico], [Lado Fisico do Bico], [Ponto de Atendimento da Bomba], [Lado Fisico do PA], [Preco Venda Prazo] FROM Bomba" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBomba = New adodb.Recordset
    rsBomba.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBomba.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBomba.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsBomba("Empresa").Value
    Codigo = rsBomba("Codigo").Value
    TipoCombustivel = rsBomba("Tipo de Combustivel").Value
    PrecoCusto = rsBomba("Preco de Custo").Value
    PrecoVenda = rsBomba("Preco de Venda").Value
    NumeroTanque = rsBomba("Numero do Tanque").Value
    NumeroIlha = rsBomba("Numero da Ilha").Value
    TipoPreco = rsBomba("Tipo de Preco").Value
    CodigoProduto = rsBomba("Codigo do Produto").Value
    NumeroSerie = rsBomba("Numero de Serie").Value
    CodigoFisicoBomba = rsBomba("Codigo Fisico da Bomba").Value
    CodigoFisicoBico = rsBomba("Codigo Fisico do Bico").Value
    LadoFisicoBico = rsBomba("Lado Fisico do Bico").Value
    PontoAtendimentoBomba = rsBomba("Ponto de Atendimento da Bomba").Value
    LadoFisicoPA = rsBomba("Lado Fisico do PA").Value
    PrecoVendaPrazo = rsBomba("Preco Venda Prazo").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

