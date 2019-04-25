VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfiguracaoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoCheque As Integer
Private mValorSuperior As Currency
Private mValorEsquerda As Currency
Private mExtenso1Superior As Currency
Private mExtenso1Esquerda As Currency
Private mExtenso2Superior As Currency
Private mExtenso2Esquerda As Currency
Private mFavorecidoSuperior As Currency
Private mFavorecidoEsquerda As Currency
Private mCidadeSuperior As Currency
Private mCidadeEsquerda As Currency
Private mDiaSuperior As Currency
Private mDiaEsquerda As Currency
Private mMesSuperior As Currency
Private mMesEsquerda As Currency
Private mAnoSuperior As Currency
Private mAnoEsquerda As Currency

Private rsConfiguracaoCheque As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoCheque() As Integer
    TipoCheque = mTipoCheque
End Property
Public Property Let TipoCheque(ByVal Valor As Integer)
    mTipoCheque = Valor
End Property


Public Property Get ValorSuperior() As Currency
    ValorSuperior = mValorSuperior
End Property
Public Property Let ValorSuperior(ByVal Valor As Currency)
    mValorSuperior = Valor
End Property


Public Property Get ValorEsquerda() As Currency
    ValorEsquerda = mValorEsquerda
End Property
Public Property Let ValorEsquerda(ByVal Valor As Currency)
    mValorEsquerda = Valor
End Property


Public Property Get Extenso1Superior() As Currency
    Extenso1Superior = mExtenso1Superior
End Property
Public Property Let Extenso1Superior(ByVal Valor As Currency)
    mExtenso1Superior = Valor
End Property


Public Property Get Extenso1Esquerda() As Currency
    Extenso1Esquerda = mExtenso1Esquerda
End Property
Public Property Let Extenso1Esquerda(ByVal Valor As Currency)
    mExtenso1Esquerda = Valor
End Property


Public Property Get Extenso2Superior() As Currency
    Extenso2Superior = mExtenso2Superior
End Property
Public Property Let Extenso2Superior(ByVal Valor As Currency)
    mExtenso2Superior = Valor
End Property


Public Property Get Extenso2Esquerda() As Currency
    Extenso2Esquerda = mExtenso2Esquerda
End Property
Public Property Let Extenso2Esquerda(ByVal Valor As Currency)
    mExtenso2Esquerda = Valor
End Property


Public Property Get FavorecidoSuperior() As Currency
    FavorecidoSuperior = mFavorecidoSuperior
End Property
Public Property Let FavorecidoSuperior(ByVal Valor As Currency)
    mFavorecidoSuperior = Valor
End Property


Public Property Get FavorecidoEsquerda() As Currency
    FavorecidoEsquerda = mFavorecidoEsquerda
End Property
Public Property Let FavorecidoEsquerda(ByVal Valor As Currency)
    mFavorecidoEsquerda = Valor
End Property


Public Property Get CidadeSuperior() As Currency
    CidadeSuperior = mCidadeSuperior
End Property
Public Property Let CidadeSuperior(ByVal Valor As Currency)
    mCidadeSuperior = Valor
End Property


Public Property Get CidadeEsquerda() As Currency
    CidadeEsquerda = mCidadeEsquerda
End Property
Public Property Let CidadeEsquerda(ByVal Valor As Currency)
    mCidadeEsquerda = Valor
End Property


Public Property Get DiaSuperior() As Currency
    DiaSuperior = mDiaSuperior
End Property
Public Property Let DiaSuperior(ByVal Valor As Currency)
    mDiaSuperior = Valor
End Property


Public Property Get DiaEsquerda() As Currency
    DiaEsquerda = mDiaEsquerda
End Property
Public Property Let DiaEsquerda(ByVal Valor As Currency)
    mDiaEsquerda = Valor
End Property


Public Property Get MesSuperior() As Currency
    MesSuperior = mMesSuperior
End Property
Public Property Let MesSuperior(ByVal Valor As Currency)
    mMesSuperior = Valor
End Property


Public Property Get MesEsquerda() As Currency
    MesEsquerda = mMesEsquerda
End Property
Public Property Let MesEsquerda(ByVal Valor As Currency)
    mMesEsquerda = Valor
End Property


Public Property Get AnoSuperior() As Currency
    AnoSuperior = mAnoSuperior
End Property
Public Property Let AnoSuperior(ByVal Valor As Currency)
    mAnoSuperior = Valor
End Property


Public Property Get AnoEsquerda() As Currency
    AnoEsquerda = mAnoEsquerda
End Property
Public Property Let AnoEsquerda(ByVal Valor As Currency)
    mAnoEsquerda = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nTipoCheque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Tipo de Cheque] = " & nTipoCheque
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsConfiguracaoCheque.Close
    Set rsConfiguracaoCheque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO ConfiguracaoCheque ( Empresa, [Tipo de Cheque], [Valor Superior], [Valor Esquerda], [Extenso 1 Superior], "
    gSQL = gSQL & "[Extenso 1 Esquerda], [Extenso 2 Superior], [Extenso 2 Esquerda], [Favorecido Superior], [Favorecido Esquerda], "
    gSQL = gSQL & "[Cidade Superior], [Cidade Esquerda], [Dia Superior], [Dia Esquerda], [Mes Superior], "
    gSQL = gSQL & "[Mes Esquerda], [Ano Superior], [Ano Esquerda] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mTipoCheque, ", ")
    Call sqlValor(1, mValorSuperior, ", ")
    Call sqlValor(1, mValorEsquerda, ", ")
    Call sqlValor(1, mExtenso1Superior, ", ")
    Call sqlValor(1, mExtenso1Esquerda, ", ")
    Call sqlValor(1, mExtenso2Superior, ", ")
    Call sqlValor(1, mExtenso2Esquerda, ", ")
    Call sqlValor(1, mFavorecidoSuperior, ", ")
    Call sqlValor(1, mFavorecidoEsquerda, ", ")
    Call sqlValor(1, mCidadeSuperior, ", ")
    Call sqlValor(1, mCidadeEsquerda, ", ")
    Call sqlValor(1, mDiaSuperior, ", ")
    Call sqlValor(1, mDiaEsquerda, ", ")
    Call sqlValor(1, mMesSuperior, ", ")
    Call sqlValor(1, mMesEsquerda, ", ")
    Call sqlValor(1, mAnoSuperior, ", ")
    Call sqlValor(1, mAnoEsquerda, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nTipoCheque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE ConfiguracaoCheque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Tipo de Cheque] = ", mTipoCheque)
    Call sqlValor(2, ", [Valor Superior] = ", mValorSuperior)
    Call sqlValor(2, ", [Valor Esquerda] = ", mValorEsquerda)
    Call sqlValor(2, ", [Extenso 1 Superior] = ", mExtenso1Superior)
    Call sqlValor(2, ", [Extenso 1 Esquerda] = ", mExtenso1Esquerda)
    Call sqlValor(2, ", [Extenso 2 Superior] = ", mExtenso2Superior)
    Call sqlValor(2, ", [Extenso 2 Esquerda] = ", mExtenso2Esquerda)
    Call sqlValor(2, ", [Favorecido Superior] = ", mFavorecidoSuperior)
    Call sqlValor(2, ", [Favorecido Esquerda] = ", mFavorecidoEsquerda)
    Call sqlValor(2, ", [Cidade Superior] = ", mCidadeSuperior)
    Call sqlValor(2, ", [Cidade Esquerda] = ", mCidadeEsquerda)
    Call sqlValor(2, ", [Dia Superior] = ", mDiaSuperior)
    Call sqlValor(2, ", [Dia Esquerda] = ", mDiaEsquerda)
    Call sqlValor(2, ", [Mes Superior] = ", mMesSuperior)
    Call sqlValor(2, ", [Mes Esquerda] = ", mMesEsquerda)
    Call sqlValor(2, ", [Ano Superior] = ", mAnoSuperior)
    Call sqlValor(2, ", [Ano Esquerda] = ", mAnoEsquerda)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Tipo de Cheque] = " & nTipoCheque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nTipoCheque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM ConfiguracaoCheque"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Tipo de Cheque] = " & nTipoCheque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Tipo de Cheque], [Valor Superior], [Valor Esquerda], [Extenso 1 Superior], [Extenso 1 Esquerda], [Extenso 2 Superior], [Extenso 2 Esquerda], [Favorecido Superior], [Favorecido Esquerda], [Cidade Superior], [Cidade Esquerda], [Dia Superior], [Dia Esquerda], [Mes Superior], [Mes Esquerda], [Ano Superior], [Ano Esquerda] FROM ConfiguracaoCheque" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConfiguracaoCheque = New adodb.Recordset
    rsConfiguracaoCheque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConfiguracaoCheque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConfiguracaoCheque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsConfiguracaoCheque("Empresa").Value
    TipoCheque = rsConfiguracaoCheque("Tipo de Cheque").Value
    ValorSuperior = rsConfiguracaoCheque("Valor Superior").Value
    ValorEsquerda = rsConfiguracaoCheque("Valor Esquerda").Value
    Extenso1Superior = rsConfiguracaoCheque("Extenso 1 Superior").Value
    Extenso1Esquerda = rsConfiguracaoCheque("Extenso 1 Esquerda").Value
    Extenso2Superior = rsConfiguracaoCheque("Extenso 2 Superior").Value
    Extenso2Esquerda = rsConfiguracaoCheque("Extenso 2 Esquerda").Value
    FavorecidoSuperior = rsConfiguracaoCheque("Favorecido Superior").Value
    FavorecidoEsquerda = rsConfiguracaoCheque("Favorecido Esquerda").Value
    CidadeSuperior = rsConfiguracaoCheque("Cidade Superior").Value
    CidadeEsquerda = rsConfiguracaoCheque("Cidade Esquerda").Value
    DiaSuperior = rsConfiguracaoCheque("Dia Superior").Value
    DiaEsquerda = rsConfiguracaoCheque("Dia Esquerda").Value
    MesSuperior = rsConfiguracaoCheque("Mes Superior").Value
    MesEsquerda = rsConfiguracaoCheque("Mes Esquerda").Value
    AnoSuperior = rsConfiguracaoCheque("Ano Superior").Value
    AnoEsquerda = rsConfiguracaoCheque("Ano Esquerda").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

