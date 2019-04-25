VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mEndereco As String
Private mBairro As String
Private mCidade As String
Private mEstado As String
Private mCEP As String
Private mCGC As String
Private mInscricaoEstadual As String
Private mDistribuidora As String
Private mInativo As Boolean
Private mResponsavelLegal As String
Private mEmpresaGlobal As Integer
Private mTipoEmpresaPosto As Boolean
Private mTipoEmpresaConveniencia As Boolean
Private mTelefone As String

Private rsEmpresa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Endereco() As String
    Endereco = mEndereco
End Property
Public Property Let Endereco(ByVal Valor As String)
    mEndereco = Valor
End Property


Public Property Get Bairro() As String
    Bairro = mBairro
End Property
Public Property Let Bairro(ByVal Valor As String)
    mBairro = Valor
End Property


Public Property Get Cidade() As String
    Cidade = mCidade
End Property
Public Property Let Cidade(ByVal Valor As String)
    mCidade = Valor
End Property


Public Property Get Estado() As String
    Estado = mEstado
End Property
Public Property Let Estado(ByVal Valor As String)
    mEstado = Valor
End Property


Public Property Get CEP() As String
    CEP = mCEP
End Property
Public Property Let CEP(ByVal Valor As String)
    mCEP = Valor
End Property


Public Property Get CGC() As String
    CGC = mCGC
End Property
Public Property Let CGC(ByVal Valor As String)
    mCGC = Valor
End Property


Public Property Get InscricaoEstadual() As String
    InscricaoEstadual = mInscricaoEstadual
End Property
Public Property Let InscricaoEstadual(ByVal Valor As String)
    mInscricaoEstadual = Valor
End Property


Public Property Get Distribuidora() As String
    Distribuidora = mDistribuidora
End Property
Public Property Let Distribuidora(ByVal Valor As String)
    mDistribuidora = Valor
End Property


Public Property Get Inativo() As Boolean
    Inativo = mInativo
End Property
Public Property Let Inativo(ByVal Valor As Boolean)
    mInativo = Valor
End Property


Public Property Get ResponsavelLegal() As String
    ResponsavelLegal = mResponsavelLegal
End Property
Public Property Let ResponsavelLegal(ByVal Valor As String)
    mResponsavelLegal = Valor
End Property


Public Property Get EmpresaGlobal() As Integer
    EmpresaGlobal = mEmpresaGlobal
End Property
Public Property Let EmpresaGlobal(ByVal Valor As Integer)
    mEmpresaGlobal = Valor
End Property


Public Property Get TipoEmpresaPosto() As Boolean
    TipoEmpresaPosto = mTipoEmpresaPosto
End Property
Public Property Let TipoEmpresaPosto(ByVal Valor As Boolean)
    mTipoEmpresaPosto = Valor
End Property


Public Property Get TipoEmpresaConveniencia() As Boolean
    TipoEmpresaConveniencia = mTipoEmpresaConveniencia
End Property
Public Property Let TipoEmpresaConveniencia(ByVal Valor As Boolean)
    mTipoEmpresaConveniencia = Valor
End Property


Public Property Get Telefone() As String
    Telefone = mTelefone
End Property
Public Property Let Telefone(ByVal Valor As String)
    mTelefone = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
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
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
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
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Empresas ( Codigo, Nome, Endereco, Bairro, Cidade, "
    gSQL = gSQL & "Estado, CEP, Cgc, [Inscricao Estadual], Distribuidora, "
    gSQL = gSQL & "Inativo, [Responsavel Legal], [Empresa Global], "
    gSQL = gSQL & "[Tipo de Empresa Posto], [Tipo de Empresa Conveniencia], "
    gSQL = gSQL & "Telefone ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mEndereco, ", ")
    Call sqlTexto(1, mBairro, ", ")
    Call sqlTexto(1, mCidade, ", ")
    Call sqlTexto(1, mEstado, ", ")
    Call sqlTexto(1, mCEP, ", ")
    Call sqlTexto(1, mCGC, ", ")
    Call sqlTexto(1, mInscricaoEstadual, ", ")
    Call sqlTexto(1, mDistribuidora, ", ")
    Call sqlBoolean(1, mInativo, ", ")
    Call sqlTexto(1, mResponsavelLegal, ", ")
    Call sqlNumero(1, mEmpresaGlobal, ", ")
    Call sqlBoolean(1, mTipoEmpresaPosto, ", ")
    Call sqlBoolean(1, mTipoEmpresaConveniencia, ", ")
    Call sqlTexto(1, mTelefone, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Empresas SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", Endereco = ", mEndereco)
    Call sqlTexto(2, ", Bairro = ", mBairro)
    Call sqlTexto(2, ", Cidade = ", mCidade)
    Call sqlTexto(2, ", Estado = ", mEstado)
    Call sqlTexto(2, ", CEP = ", mCEP)
    Call sqlTexto(2, ", Cgc = ", mCGC)
    Call sqlTexto(2, ", [Inscricao Estadual] = ", mInscricaoEstadual)
    Call sqlTexto(2, ", Distribuidora = ", mDistribuidora)
    Call sqlBoolean(2, ", Inativo = ", mInativo)
    Call sqlTexto(2, ", [Responsavel Legal] = ", mResponsavelLegal)
    Call sqlNumero(2, ", [Empresa Global] = ", mEmpresaGlobal)
    Call sqlBoolean(2, ", [Tipo de Empresa Posto] = ", mTipoEmpresaPosto)
    Call sqlBoolean(2, ", [Tipo de Empresa Conveniencia] = ", mTipoEmpresaConveniencia)
    Call sqlTexto(2, ", Telefone = ", mTelefone)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Empresas"
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo,  Nome,  Endereco,  Bairro,  Cidade,  Estado,  CEP,  Cgc,  [Inscricao Estadual],  Distribuidora,  Inativo, [Responsavel Legal], [Empresa Global], [Tipo de Empresa Posto], [Tipo de Empresa Conveniencia], Telefone FROM Empresas" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEmpresa = New adodb.Recordset
    rsEmpresa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEmpresa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEmpresa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsEmpresa("Codigo").Value
    Nome = rsEmpresa("Nome").Value
    Endereco = rsEmpresa("Endereco").Value
    Bairro = rsEmpresa("Bairro").Value
    Cidade = rsEmpresa("Cidade").Value
    Estado = rsEmpresa("Estado").Value
    CEP = rsEmpresa("CEP").Value
    CGC = rsEmpresa("Cgc").Value
    InscricaoEstadual = rsEmpresa("Inscricao Estadual").Value
    Distribuidora = rsEmpresa("Distribuidora").Value
    Inativo = rsEmpresa("Inativo").Value
    ResponsavelLegal = rsEmpresa("Responsavel Legal").Value
    EmpresaGlobal = rsEmpresa("Empresa Global").Value
    TipoEmpresaPosto = rsEmpresa("Tipo de Empresa Posto").Value
    TipoEmpresaConveniencia = rsEmpresa("Tipo de Empresa Conveniencia").Value
    Telefone = rsEmpresa("Telefone").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsEmpresa.Close
    Set rsEmpresa = Nothing
End Function

