VERSION 5.00
Begin VB.Form indexa 
   Caption         =   "Indexador de Banco de Dados"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_inicio2 
      Caption         =   "I&nício 2"
      Height          =   435
      Left            =   660
      TabIndex        =   1
      ToolTipText     =   "Cria [Valor do Desconto], Nome, [CPF CNPJ] e atribui valor 0 para [Valor do Desconto]"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmd_inicio 
      Caption         =   "&Início 1"
      Height          =   435
      Left            =   660
      TabIndex        =   0
      ToolTipText     =   "Cliente 0, Operador = 1, Reindexa"
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "indexa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl_dados As Table
Dim tbl_cliente As Table
Dim tbl_produto As Table
Dim tbl_movimento_cupom_fiscal As Table
Private Sub CriaCamposTabelaMovimentoCupomFiscal()
    On Error GoTo FileError
    'Cria na tabela de Movimento_Cupom_Fiscal
    'Valor do Desconto
    'Valor do Acrescimo
    Dim gdbCurrentDB As Database
    Dim gtdfTableDef As TableDef
    Dim fld As Field
    Set gdbCurrentDB = OpenDatabase("C:\VB5\SGP\DATA\SGP_DATA.MDB")
    Set gtdfTableDef = gdbCurrentDB.TableDefs(StripConnect("Movimento_Cupom_Fiscal"))
    Set fld = gtdfTableDef.CreateField()
    With fld
        .Name = "Valor do Desconto"
        .Type = 5
        '.Size = 8
        .OrdinalPosition = 22
        '.CollatingOrder = 1033
    End With
    gtdfTableDef.Fields.Append fld
    
    Set fld = gtdfTableDef.CreateField()
    With fld
        .Name = "Nome"
        .Type = 10
        .OrdinalPosition = 23
        .Size = 40
        '.CollatingOrder = 1033
    End With
    gtdfTableDef.Fields.Append fld
    
    Set fld = gtdfTableDef.CreateField()
    With fld
        .Name = "CPF CNPJ"
        .Type = 10
        .Size = 20
        .OrdinalPosition = 24
        '.CollatingOrder = 1033
    End With
    gtdfTableDef.Fields.Append fld
    
    gdbCurrentDB.Execute "Update Movimento_Cupom_Fiscal Set [Valor do Desconto] = 0"
    gdbCurrentDB.Close
    Exit Sub
FileError:
    MsgBox Error & "Não foi possível criar campos na tabela de Movimento_Cupom_Fiscal"
    Exit Sub
End Sub
Private Sub Finaliza()
    tbl_dados.Close
    tbl_cliente.Close
    tbl_produto.Close
    bd_sgp.Close
    End
End Sub
Function StripConnect(rsTblName As String) As String
  If InStr(rsTblName, "->") > 0 Then
    StripConnect = Left(rsTblName, InStr(rsTblName, "->") - 2)
  Else
    StripConnect = rsTblName
  End If
End Function
Private Sub cmd_inicio_Click()
    Dim i As Integer
    Dim i2 As Integer
    Set tbl_movimento_cupom_fiscal = bd_sgp.OpenTable("Movimento_Cupom_Fiscal")
    tbl_dados.MoveFirst
    tbl_dados.Edit
    tbl_dados![Empresa 2] = 0
    tbl_dados.Update
    tbl_dados.MoveFirst
    With tbl_cliente
        .AddNew
        !Codigo = 0
        ![Razao Social] = "."
        !Endereco = "."
        !Bairro = "."
        !Cidade = "."
        !UF = "."
        !CEP = "74000000"
        !Telefone = "062"
        !CPF = "1"
        !Cgc = "1"
        !Identidade = "1"
        ![Orgao Emissor] = "."
        ![Inscricao Estadual] = "1"
        ![Codigo do Vencimento] = 1
        ![Codigo do Convenio] = 1
        !Empresa = 1
        !Inativo = False
        ![Imprime Duplicata] = False
        .Update
    End With
    bd_sgp.Execute "Update Movimento_Cupom_Fiscal Set Operador = 1"
    tbl_movimento_cupom_fiscal.Close
    i = 0
    For i2 = 0 To 32000
        tbl_produto.MoveFirst
        Do Until tbl_produto.EOF
            i = i + 1
            tbl_produto.MoveNext
        Loop
    Next
End Sub
Private Sub cmd_inicio2_Click()
    CriaCamposTabelaMovimentoCupomFiscal
End Sub
Private Sub Form_Load()
    Call ChamaDrive
    ChDir "\VB5\SGP\DATA"
    Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
    Set tbl_cliente = bd_sgp.OpenTable("Cliente")
    Set tbl_dados = bd_sgp.OpenTable("dados")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    tbl_produto.Index = "id_codigo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
