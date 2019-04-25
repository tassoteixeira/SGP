VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfiguracaoDiversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNome As String
Private mCampoUtilizado As String
Private mQuantidadeCasaDecimal As Integer
Private mQuantidadeCaracter As Integer
Private mData As Date
Private mHora As Date
Private mCodigo As Long
Private mValor As Currency
Private mVerdadeiro As Boolean
Private mTexto As String

Private rsConfiguracaoDiversa As New adodb.Recordset
Private lMovDocumentoEletronicoCabecalho As New cMovDocEletronicoCabecalho


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get CampoUtilizado() As String
    CampoUtilizado = mCampoUtilizado
End Property
Public Property Let CampoUtilizado(ByVal Valor As String)
    mCampoUtilizado = Valor
End Property


Public Property Get QuantidadeCasaDecimal() As Integer
    QuantidadeCasaDecimal = mQuantidadeCasaDecimal
End Property
Public Property Let QuantidadeCasaDecimal(ByVal Valor As Integer)
    mQuantidadeCasaDecimal = Valor
End Property


Public Property Get QuantidadeCaracter() As Integer
    QuantidadeCaracter = mQuantidadeCaracter
End Property
Public Property Let QuantidadeCaracter(ByVal Valor As Integer)
    mQuantidadeCaracter = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Hora() As Date
    Hora = mHora
End Property
Public Property Let Hora(ByVal Valor As Date)
    mHora = Valor
End Property


Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Long)
    mCodigo = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get Verdadeiro() As Boolean
    Verdadeiro = mVerdadeiro
End Property
Public Property Let Verdadeiro(ByVal Valor As Boolean)
    mVerdadeiro = Valor
End Property


Public Property Get Texto() As String
    Texto = mTexto
End Property
Public Property Let Texto(ByVal Valor As String)
    mTexto = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function BuscaProximoCodigo(ByVal pEmpresa As Integer, ByVal pNome As String, ByVal pVerificaExistenciaNFCe As Boolean) As String
Dim xCondicao As String
Dim xRetornoBusca As Boolean
Dim lRecordsAffected As Long
Dim xProximoCodigo As Long
Dim xSerie As String
Dim xEncontrou As Boolean
Dim xContador As Integer

On Error GoTo trata_erro

    xRetornoBusca = False
    BuscaProximoCodigo = ""
    xProximoCodigo = 0
    xSerie = ""
    xContador = 1
    xEncontrou = False
    
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNome)
    gSQL = PreparaSQL(xCondicao, "")
    xRetornoBusca = Localizar(1)
    If xRetornoBusca = False Then
        rsConfiguracaoDiversa.Close
        Call CriaRegistro(pEmpresa, pNome, "Codigo", 0)
        gSQL = PreparaSQL(xCondicao, "")
        xRetornoBusca = Localizar(1)
    End If
    If xRetornoBusca = True Then
        
        'BuscaProximoCodigo = (rsConfiguracaoDiversa("Codigo").Value + 1) & "|@|" & rsConfiguracaoDiversa("Texto").Value & "|@|"
        xProximoCodigo = CLng(rsConfiguracaoDiversa("Codigo").Value + 1)
        xSerie = rsConfiguracaoDiversa("Texto").Value
        
        If pVerificaExistenciaNFCe Then
            Do Until xEncontrou = True
                If Not VerificaSeExisteNumeroNFCe(xProximoCodigo, xSerie, pEmpresa) Then
                    xEncontrou = True
                Else
                    xProximoCodigo = xProximoCodigo + 1
                    xContador = xContador + 1
                End If
            Loop
        End If
        
        BuscaProximoCodigo = xProximoCodigo & "|@|" & xSerie & "|@|" '(rsConfiguracaoDiversa("Codigo").Value + 1) & "|@|" & rsConfiguracaoDiversa("Texto").Value & "|@|"
        
        'Altera o Registro
        gSQL = "UPDATE ConfiguracaoDiversa SET"
        'gSQL = gSQL & " Codigo = Codigo + 1"
        gSQL = gSQL & " Codigo = Codigo + " & xContador
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        gSQL = gSQL & " AND Nome = " & preparaTexto(pNome)
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected = 0 Then
            Call CriaLogCadastroDll2("cConfiguracaoDiversa:BuscaProximoCodigo - Erro ao alterar codigo.", Err.Description, gSQL)
        End If
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    Call CriaLogCadastroDll2("cConfiguracaoDiversa:BuscaProximoCodigo - Erro desconhecido.", Err.Description, gSQL)
End Function
Private Function VerificaSeExisteNumeroNFCe(ByVal pNumero As String, pSerie As String, ByVal pCodigoEmpresa As Integer) As Boolean

    VerificaSeExisteNumeroNFCe = lMovDocumentoEletronicoCabecalho.ExisteDocumentoEletronico(pCodigoEmpresa, False, True, "65", pSerie, pNumero)


End Function

Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    Else
        LocalizarCodigo = False
        xCondicao = " WHERE Empresa = 1"
        xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNome)
        gSQL = PreparaSQL(xCondicao, "")
        LocalizarCodigo = Localizar(1)
        If LocalizarCodigo = True Then
            AtribuiValor
        End If
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Nome < " & preparaTexto(mNome)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Nome DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Nome ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Nome > " & preparaTexto(mNome)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Nome ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Nome DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsConfiguracaoDiversa.Close
    Set rsConfiguracaoDiversa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO ConfiguracaoDiversa ( Empresa, Nome, [Campo Utilizado], [Quantidade de Casa Decimal], [Quantidade de Caracter], "
    gSQL = gSQL & "Data, Hora, Codigo, Valor, Verdadeiro, "
    gSQL = gSQL & "Texto ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mCampoUtilizado, ", ")
    Call sqlNumero(1, mQuantidadeCasaDecimal, ", ")
    Call sqlNumero(1, mQuantidadeCaracter, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlData(1, mHora, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlBoolean(1, mVerdadeiro, ", ")
    Call sqlTexto(1, mTexto, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pNome As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE ConfiguracaoDiversa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", [Campo Utilizado] = ", mCampoUtilizado)
    Call sqlNumero(2, ", [Quantidade de Casa Decimal] = ", mQuantidadeCasaDecimal)
    Call sqlNumero(2, ", [Quantidade de Caracter] = ", mQuantidadeCaracter)
    Call sqlData(2, ", Data = ", mData)
    Call sqlData(2, ", Hora = ", mHora)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlBoolean(2, ", Verdadeiro = ", mVerdadeiro)
    Call sqlTexto(2, ", Texto = ", mTexto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Nome = " & preparaTexto(pNome)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pNome As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM ConfiguracaoDiversa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Nome = " & preparaTexto(pNome)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function CriaRegistro(ByVal pEmpresa As Integer, ByVal pNome As String, ByVal pCampoUtilizado As String, ByVal pValor As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    CriaRegistro = False
    gSQL = "INSERT INTO ConfiguracaoDiversa ( Empresa, Nome, [Campo Utilizado], [Quantidade de Casa Decimal], [Quantidade de Caracter], "
    gSQL = gSQL & "Data, Hora, Codigo, Valor, Verdadeiro, "
    gSQL = gSQL & "Texto ) VALUES ( "
    Call sqlNumero(1, pEmpresa, ", ")
    Call sqlTexto(1, pNome, ", ")
    Call sqlTexto(1, pCampoUtilizado, ", ")
    Call sqlNumero(1, 0, ", ")
    Call sqlNumero(1, 0, ", ")
    Call sqlData(1, "00:00:00", ", ")
    Call sqlData(1, "00:00:00", ", ")
    If pCampoUtilizado = "Codigo" Then
        Call sqlNumero(1, CLng(pValor), ", ")
    Else
        Call sqlNumero(1, 0, ", ")
    End If
    Call sqlValor(1, 0, ", ")
    Call sqlBoolean(1, False, ", ")
    If pNome = "NFCe: Numero" Then
        Call sqlTexto(1, "1", " )")
    Else
        Call sqlTexto(1, "", " )")
    End If
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        CriaRegistro = True
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("cConfiguracaoDiversa:CriaRegistro - Erro desconhecido.", Err.Description, gSQL)
End Function
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Nome, [Campo Utilizado], [Quantidade de Casa Decimal], [Quantidade de Caracter], Data, Hora, Codigo, Valor, Verdadeiro, Texto FROM ConfiguracaoDiversa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConfiguracaoDiversa = New adodb.Recordset
    rsConfiguracaoDiversa.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConfiguracaoDiversa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConfiguracaoDiversa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsConfiguracaoDiversa("Empresa").Value
    Nome = rsConfiguracaoDiversa("Nome").Value
    CampoUtilizado = rsConfiguracaoDiversa("Campo Utilizado").Value
    QuantidadeCasaDecimal = rsConfiguracaoDiversa("Quantidade de Casa Decimal").Value
    QuantidadeCaracter = rsConfiguracaoDiversa("Quantidade de Caracter").Value
    If IsNull(rsConfiguracaoDiversa("Data").Value) Then
        Data = "00:00:00"
    Else
        Data = rsConfiguracaoDiversa("Data").Value
    End If
    If IsNull(rsConfiguracaoDiversa("Hora").Value) Then
        Hora = "00:00:00"
    Else
        Hora = rsConfiguracaoDiversa("Hora").Value
    End If
    Codigo = rsConfiguracaoDiversa("Codigo").Value
    Valor = rsConfiguracaoDiversa("Valor").Value
    Verdadeiro = rsConfiguracaoDiversa("Verdadeiro").Value
    Texto = rsConfiguracaoDiversa("Texto").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

