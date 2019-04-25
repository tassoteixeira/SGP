VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form movimento_pedido_combustivel 
   Caption         =   "Pedido de Combustíveis"
   ClientHeight    =   6195
   ClientLeft      =   2340
   ClientTop       =   1425
   ClientWidth     =   7335
   Icon            =   "ped_combustivel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ped_combustivel.frx":030A
   ScaleHeight     =   6195
   ScaleWidth      =   7335
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2760
      Picture         =   "ped_combustivel.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5220
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "ped_combustivel.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5220
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "ped_combustivel.frx":3474
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Altera o registro atual."
      Top             =   5220
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "ped_combustivel.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Cria um novo registro."
      Top             =   5220
      Width           =   795
   End
   Begin TabDlg.SSTab tab_dados 
      Height          =   4995
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8811
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   529
      Enabled         =   0   'False
      TabCaption(0)   =   "Dados do Pedido"
      TabPicture(0)   =   "ped_combustivel.frx":6000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "msk_data"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "msk_data_entrega"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmFrame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_pedido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txt_pedido 
         Height          =   285
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   46
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Frame frmFrame1 
         BackColor       =   &H00C0C0C0&
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   6975
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   0
            Left            =   3060
            TabIndex        =   13
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   0
            Left            =   4380
            TabIndex        =   14
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   18
            Top             =   900
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   1
            Left            =   4380
            TabIndex        =   19
            Top             =   900
            Width           =   1095
         End
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   2
            Left            =   3060
            TabIndex        =   23
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   2
            Left            =   4380
            TabIndex        =   24
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   3
            Left            =   3060
            TabIndex        =   28
            Top             =   1620
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   3
            Left            =   4380
            TabIndex        =   29
            Top             =   1620
            Width           =   1095
         End
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   4
            Left            =   3060
            TabIndex        =   33
            Top             =   1980
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   4
            Left            =   4380
            TabIndex        =   34
            Top             =   1980
            Width           =   1095
         End
         Begin VB.TextBox txt_preco_unitario 
            Height          =   285
            Index           =   5
            Left            =   3060
            TabIndex        =   38
            Top             =   2340
            Width           =   1095
         End
         Begin VB.TextBox txt_quantidade_combustivel 
            Height          =   285
            Index           =   5
            Left            =   4380
            TabIndex        =   39
            Top             =   2340
            Width           =   1095
         End
         Begin VB.TextBox txt_total_quantidade_combustivel 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4380
            TabIndex        =   43
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Line Line10 
            X1              =   60
            X2              =   6900
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Top             =   540
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   -60
            X2              =   6780
            Y1              =   -1380
            Y2              =   -1380
         End
         Begin VB.Line Line2 
            X1              =   60
            X2              =   6900
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            X1              =   60
            X2              =   60
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Line Line5 
            X1              =   1620
            X2              =   1620
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Line Line6 
            X1              =   4260
            X2              =   4260
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Line Line7 
            X1              =   60
            X2              =   6900
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label Label11 
            Caption         =   "Produtos"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Preço Unitário"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3120
            TabIndex        =   8
            Top             =   180
            Width           =   1095
         End
         Begin VB.Line Line8 
            X1              =   5580
            X2              =   5580
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Label Label13 
            Caption         =   "Quantidade"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   4500
            TabIndex        =   9
            Top             =   180
            Width           =   915
         End
         Begin VB.Line Line9 
            X1              =   6900
            X2              =   6900
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   5700
            TabIndex        =   15
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Total"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6000
            TabIndex        =   10
            Top             =   180
            Width           =   555
         End
         Begin VB.Line Line11 
            X1              =   2940
            X2              =   2940
            Y1              =   120
            Y2              =   3120
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   1740
            TabIndex        =   12
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Estoque"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1980
            TabIndex        =   7
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "AA"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   16
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   1740
            TabIndex        =   17
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   5700
            TabIndex        =   20
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "D"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   21
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   1740
            TabIndex        =   22
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   5700
            TabIndex        =   25
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "DA"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   26
            Top             =   1620
            Width           =   1335
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   1740
            TabIndex        =   27
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   5700
            TabIndex        =   30
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "G"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   180
            TabIndex        =   31
            Top             =   1980
            Width           =   1335
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   1740
            TabIndex        =   32
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   5700
            TabIndex        =   35
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label lbl_nome_combustivel 
            Caption         =   "GA"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   180
            TabIndex        =   36
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label lbl_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   5
            Left            =   1740
            TabIndex        =   37
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Label lbl_total_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   5
            Left            =   5700
            TabIndex        =   40
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Line Line4 
            X1              =   60
            X2              =   6900
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Label Label10 
            Caption         =   "Total"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   420
            TabIndex        =   41
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label lbl_total_estoque_combustivel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1740
            TabIndex        =   42
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label lbl_total_valor_pedido 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5700
            TabIndex        =   44
            Top             =   2760
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox msk_data_entrega 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   960
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   540
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Data P/ Entrega"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Data do Pedido"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Número do Pedido"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   4560
         Width           =   1395
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5100
      TabIndex        =   53
      Top             =   5100
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "ped_combustivel.frx":601C
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "ped_combustivel.frx":759E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "ped_combustivel.frx":8A10
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "ped_combustivel.frx":9F0A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5580
      Picture         =   "ped_combustivel.frx":B404
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5220
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6480
      Picture         =   "ped_combustivel.frx":CA0E
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5220
      Width           =   795
   End
End
Attribute VB_Name = "movimento_pedido_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_pedido_combustivel As Integer
Dim lOpcao As Integer
Dim l_data As Date
Dim l_pedido As String
Dim tbl_bomba As Table
Dim tbl_combustivel As Table
Dim tbl_pedido_combustivel As Table
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    l_data = msk_data
    l_pedido = txt_pedido
    With tbl_pedido_combustivel
        !Empresa = g_empresa
        ![Data do Pedido] = msk_data
        !numero = txt_pedido
        ![Data para Entrega] = msk_data_entrega
        ![Preco de Custo A] = fValidaValor4(txt_preco_unitario(0))
        ![Quantidade A] = fValidaValor1(txt_quantidade_combustivel(0))
        ![Estoque Anterior A] = fValidaValor1(lbl_estoque_combustivel(0).Caption)
        ![Preco de Custo AA] = fValidaValor4(txt_preco_unitario(1))
        ![Quantidade AA] = fValidaValor1(txt_quantidade_combustivel(1))
        ![Estoque Anterior AA] = fValidaValor1(lbl_estoque_combustivel(1).Caption)
        ![Preco de Custo D] = fValidaValor4(txt_preco_unitario(2))
        ![Quantidade D] = fValidaValor1(txt_quantidade_combustivel(2))
        ![Estoque Anterior D] = fValidaValor1(lbl_estoque_combustivel(2).Caption)
        ![Preco de Custo DA] = fValidaValor4(txt_preco_unitario(3))
        ![Quantidade DA] = fValidaValor1(txt_quantidade_combustivel(3))
        ![Estoque Anterior DA] = fValidaValor1(lbl_estoque_combustivel(3).Caption)
        ![Preco de Custo G] = fValidaValor4(txt_preco_unitario(4))
        ![Quantidade G] = fValidaValor1(txt_quantidade_combustivel(4))
        ![Estoque Anterior G] = fValidaValor1(lbl_estoque_combustivel(4).Caption)
        ![Preco de Custo GA] = fValidaValor4(txt_preco_unitario(5))
        ![Quantidade GA] = fValidaValor1(txt_quantidade_combustivel(5))
        ![Estoque Anterior GA] = fValidaValor1(lbl_estoque_combustivel(5).Caption)
    End With
End Sub
Private Sub AtualTela()
    With tbl_pedido_combustivel
        l_data = ![Data do Pedido]
        l_pedido = !numero
        msk_data = ![Data do Pedido]
        txt_pedido = !numero
        msk_data_entrega = ![Data para Entrega]
        lbl_estoque_combustivel(0).Caption = Format(![Estoque Anterior A], "###,##0.0")
        txt_preco_unitario(0) = Format(![Preco de Custo A], "###,##0.0000")
        txt_quantidade_combustivel(0) = Format(![Quantidade A], "###,##0.0")
        lbl_total_combustivel(0).Caption = Format((![Quantidade A] * ![Preco de Custo A]), "###,##0.0")
        lbl_estoque_combustivel(1).Caption = Format(![Estoque Anterior AA], "###,##0.0")
        txt_preco_unitario(1) = Format(![Preco de Custo AA], "###,##0.0000")
        txt_quantidade_combustivel(1) = Format(![Quantidade AA], "###,##0.0")
        lbl_total_combustivel(1).Caption = Format((![Quantidade AA] * ![Preco de Custo AA]), "###,##0.0")
        lbl_estoque_combustivel(2).Caption = Format(![Estoque Anterior D], "###,##0.0")
        txt_preco_unitario(2) = Format(![Preco de Custo D], "###,##0.0000")
        txt_quantidade_combustivel(2) = Format(![Quantidade D], "###,##0.0")
        lbl_total_combustivel(2).Caption = Format((![Quantidade D] * ![Preco de Custo D]), "###,##0.0")
        lbl_estoque_combustivel(3).Caption = Format(![Estoque Anterior DA], "###,##0.0")
        txt_preco_unitario(3) = Format(![Preco de Custo DA], "###,##0.0000")
        txt_quantidade_combustivel(3) = Format(![Quantidade DA], "###,##0.0")
        lbl_total_combustivel(3).Caption = Format((![Quantidade DA] * ![Preco de Custo DA]), "###,##0.0")
        lbl_estoque_combustivel(4).Caption = Format(![Estoque Anterior G], "###,##0.0")
        txt_preco_unitario(4) = Format(![Preco de Custo G], "###,##0.0000")
        txt_quantidade_combustivel(4) = Format(![Quantidade G], "###,##0.0")
        lbl_total_combustivel(4).Caption = Format((![Quantidade G] * ![Preco de Custo G]), "###,##0.0")
        lbl_estoque_combustivel(5).Caption = Format(![Estoque Anterior GA], "###,##0.0")
        txt_preco_unitario(5) = Format(![Preco de Custo GA], "###,##0.0000")
        txt_quantidade_combustivel(5) = Format(![Quantidade GA], "###,##0.0")
        lbl_total_combustivel(5).Caption = Format((![Quantidade GA] * ![Preco de Custo GA]), "###,##0.0")
        CalculaTotais
    End With
    tab_dados.Enabled = True
End Sub
Private Sub CalculaTotais()
    Dim x_total As Currency
    Dim i As Integer
    x_total = 0
    For i = 0 To 5
        x_total = x_total + fValidaValor1(lbl_total_combustivel(i))
    Next
    lbl_total_valor_pedido = Format(x_total, "###,##0.00")
    x_total = 0
    For i = 0 To 5
        x_total = x_total + fValidaValor1(txt_quantidade_combustivel(i))
    Next
    txt_total_quantidade_combustivel = Format(x_total, "###,##0.0")
    x_total = 0
    For i = 0 To 5
        x_total = x_total + fValidaValor1(lbl_estoque_combustivel(i))
    Next
    lbl_total_estoque_combustivel = Format(x_total, "###,##0.0")
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    If tbl_pedido_combustivel.RecordCount > 0 Then
        tbl_pedido_combustivel.Seek "<", g_empresa, CDate("31/12/2500"), "ZZZZZZZZZZ"
        If Not tbl_pedido_combustivel.NoMatch Then
            If tbl_pedido_combustivel!Empresa = g_empresa Then
                AtualTela
                BuscaDados = True
            End If
        End If
    End If
End Function
Function BuscaRegistro(x_data As Date, x_pedido As String) As Boolean
    BuscaRegistro = False
    tbl_pedido_combustivel.Seek "=", g_empresa, x_data, x_pedido
    If Not tbl_pedido_combustivel.NoMatch Then
        BuscaRegistro = True
        AtualTela
    End If
End Function
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    tbl_bomba.Close
    tbl_combustivel.Close
    tbl_pedido_combustivel.Close
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    tab_dados.Enabled = True
    txt_pedido.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If tbl_pedido_combustivel.RecordCount > 0 Then
        tbl_pedido_combustivel.MovePrevious
        If tbl_pedido_combustivel.BOF Then
            MsgBox "Início de Arquivo.", 48, "Atenção!"
            tbl_pedido_combustivel.MoveNext
            cmd_proximo.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaDados Then
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    Dim i As Integer
    msk_data = "__/__/____"
    txt_pedido = ""
    msk_data_entrega = "__/__/____"
    For i = 0 To 5
        lbl_estoque_combustivel(i) = ""
        txt_preco_unitario(i) = ""
        txt_quantidade_combustivel(i) = ""
        lbl_total_combustivel(i) = ""
    Next
    lbl_total_estoque_combustivel = ""
    txt_total_quantidade_combustivel = ""
    lbl_total_valor_pedido = ""
    PreencheCombustiveis
End Sub
Private Sub cmd_excluir_Click()
    If tbl_pedido_combustivel![Data do Pedido] <> "" Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            tbl_pedido_combustivel.Edit
            tbl_pedido_combustivel.Delete
            LimpaTela
            If Not BuscaDados Then
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    tab_dados.Enabled = True
    msk_data.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            tbl_pedido_combustivel.AddNew
            AtualTabe
            tbl_pedido_combustivel.Update
        ElseIf lOpcao = 2 Then
            tbl_pedido_combustivel.Edit
            AtualTabe
            tbl_pedido_combustivel.Update
        End If
        lOpcao = 0
        Call BuscaRegistro(l_data, l_pedido)
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_pedido_combustivel.name, "Pedido de Combustívelo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data do pedido.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf txt_pedido = "" Then
        MsgBox "Informe o número do pedido.", 64, "Atenção!"
        txt_pedido.SetFocus
    ElseIf Not IsDate(msk_data_entrega) Then
        MsgBox "Informe a data para entrega.", 64, "Atenção!"
        msk_data_entrega.SetFocus
    ElseIf (fValidaValor1(txt_quantidade_combustivel(0)) + fValidaValor1(txt_quantidade_combustivel(1)) + fValidaValor1(txt_quantidade_combustivel(2)) + fValidaValor1(txt_quantidade_combustivel(3)) + fValidaValor1(txt_quantidade_combustivel(4)) + fValidaValor1(txt_quantidade_combustivel(5))) = 0 Then
        MsgBox "Pedido não pode ser feito sem quantidade.", 64, "Atenção!"
        txt_quantidade_combustivel(0).SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If tbl_pedido_combustivel.RecordCount > 0 Then
        tbl_pedido_combustivel.Seek ">", g_empresa, CDate("01/01/1900"), " "
        If Not tbl_pedido_combustivel.NoMatch Then
            If tbl_pedido_combustivel!Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If tbl_pedido_combustivel.RecordCount > 0 Then
        tbl_pedido_combustivel.MoveNext
        If tbl_pedido_combustivel.EOF Then
            MsgBox "fIM de Arquivo.", 48, "Atenção!"
            tbl_pedido_combustivel.MovePrevious
            cmd_anterior.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If tbl_pedido_combustivel.RecordCount > 0 Then
        tbl_pedido_combustivel.Seek "<", g_empresa, CDate("31/12/2500"), "ZZZZZZZZZZ"
        If Not tbl_pedido_combustivel.NoMatch Then
            If tbl_pedido_combustivel!Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    PreencheCombustiveis
    If flag_pedido_combustivel = 0 Then
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
           cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_pedido_combustivel = 0
    End If
End Sub
Private Sub Form_Deactivate()
    flag_pedido_combustivel = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_bomba = bd_sgp.OpenTable("Bomba")
    Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    Set tbl_pedido_combustivel = bd_sgp.OpenTable("Pedido_Combustivel")
    tbl_pedido_combustivel.Index = "id_data"
End Sub
Private Sub PreencheCombustiveis()
    Dim i As Integer
    tbl_bomba.Index = "id_combustivel"
    tbl_combustivel.Index = "id_codigo"
    tbl_combustivel.MoveFirst
    Do Until tbl_combustivel.EOF
        lbl_nome_combustivel(i) = tbl_combustivel!Nome
        lbl_estoque_combustivel(i) = Format(tbl_combustivel![Quantidade em Estoque], "###,##0.0")
        tbl_bomba.Seek "=", g_empresa, tbl_combustivel!Codigo
        If tbl_bomba.NoMatch Then
            txt_preco_unitario(i) = Format(0, "###,##0.0000")
        Else
            txt_preco_unitario(i) = Format(tbl_bomba!custo, "###,##0.0000")
        End If
        i = i + 1
        tbl_combustivel.MoveNext
    Loop
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_entrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_combustivel(0).SetFocus
    End If
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data) Then
        msk_data = Format(g_data_def, "dd/mm/yyyy")
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_entrega.SetFocus
    End If
End Sub
Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_preco_unitario_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_combustivel(Index).SetFocus
    End If
End Sub
Private Sub txt_preco_unitario_LostFocus(Index As Integer)
    txt_preco_unitario(Index) = Format(txt_preco_unitario(Index), "###,##0.0000")
    lbl_total_combustivel(Index).Caption = Format((Val(txt_quantidade_combustivel(Index)) * txt_preco_unitario(Index)), "###,##0.0")
    CalculaTotais
End Sub
Private Sub txt_quantidade_combustivel_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index) < 5 Then
            txt_quantidade_combustivel(Index + 1).SetFocus
        Else
            txt_pedido.SetFocus
        End If
    End If
End Sub
Private Sub txt_quantidade_combustivel_LostFocus(Index As Integer)
    txt_quantidade_combustivel(Index) = Format(txt_quantidade_combustivel(Index), "###,##0.0")
    lbl_total_combustivel(Index).Caption = Format((Val(txt_quantidade_combustivel(Index)) * txt_preco_unitario(Index)), "###,##0.0")
    CalculaTotais
End Sub
