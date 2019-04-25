VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form cadastro_tipo_documento 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cadastro de Tipo de Documentos"
   ClientHeight    =   3825
   ClientLeft      =   1890
   ClientTop       =   2805
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Bancos"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   6690
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1140
      Picture         =   "cad_tipo_documento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Confirma o registro atual."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2940
      Picture         =   "cad_tipo_documento.frx":12DA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela o registro atual."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4740
      Picture         =   "cad_tipo_documento.frx":25B4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2880
      Width           =   795
   End
   Begin VB.ListBox lbo_tipodoc 
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6480
   End
   Begin Threed.SSFrame frm_dados 
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   2160
      Width           =   6465
      _Version        =   65536
      _ExtentX        =   11404
      _ExtentY        =   1138
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   3
         Top             =   210
         Width           =   4440
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome do Documento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   1815
      End
   End
End
Attribute VB_Name = "cadastro_tipo_documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim opcao As Integer
Dim guarda
Dim tbl_tipodoc As Table
Dim l_codigo As Integer




Private Sub AtualizaLboBanco(codigo_x)
    Dim i As Integer
    For i = 0 To lbo_tipodoc.ListCount - 1
        If lbo_tipodoc.ItemData(i) = codigo_x Then
            lbo_tipodoc.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub AtualTabe()
    tbl_tipodoc!Codigo = l_codigo
    tbl_tipodoc!Nome = Format(txt_nome, "")

End Sub

Private Sub AtualTela()
    txt_nome = lbo_tipodoc.Text
    
End Sub





Private Sub PreencheLboBanco()
    lbo_tipodoc.Clear
    tbl_tipodoc.Index = "id_nome"
    If (tbl_tipodoc.BOF And tbl_tipodoc.EOF) Then
    Else
        tbl_tipodoc.MoveFirst
        Do Until tbl_tipodoc.EOF
            lbo_tipodoc.AddItem tbl_tipodoc!Nome
            lbo_tipodoc.ItemData(lbo_tipodoc.NewIndex) = tbl_tipodoc!Codigo
            tbl_tipodoc.MoveNext
        Loop
    End If
        
End Sub


Private Sub ProximoCodigo()
    tbl_tipodoc.Index = "id_codigo"
    If (tbl_tipodoc.BOF And tbl_tipodoc.EOF) Then
        l_codigo = 1
    Else
        tbl_tipodoc.MoveLast
        l_codigo = tbl_tipodoc!Codigo + 1
        g_tipodoc = l_codigo
    End If

End Sub





Private Sub cmd_cancelar_Click()
    LimpaTela
    lbo_tipodoc.SetFocus

End Sub

Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If opcao = 1 Or opcao = 3 Then
        If txt_nome.Text = "" Then
            Beep
            MsgBox "Informe o Nome do Documento!", 16, "Ajuda"
            txt_nome.SetFocus
            Exit Sub
        End If
    End If
    Select Case opcao
        Case 1
            tbl_tipodoc.AddNew
            AtualTabe
            tbl_tipodoc.Update
            g_tipodoc = l_codigo
            ProximoCodigo
            txt_nome.SetFocus
        Case 3 'ALTERAR
            tbl_tipodoc.Edit
            AtualTabe
            tbl_tipodoc.Update
            LimpaTela
            lbo_tipodoc.SetFocus
    End Select
    PreencheLboBanco
    AtualizaLboBanco l_codigo
    If opcao = 1 Then
        LimpaTela
    End If
    Exit Sub

FileError:
    ErroArquivo tbl_tipodoc.Name, "Tipo de Documentoo"
    Exit Sub

End Sub

Private Sub Finaliza()
    tbl_tipodoc.Close
    Unload Me
    
End Sub

Private Sub cmd_sair_Click()
    Finaliza

End Sub

Private Sub Form_Activate()
    LimpaTela
    ProximoCodigo
    opcao = 1
    PreencheLboBanco
    tbl_tipodoc.Index = "id_codigo"
    txt_nome.SetFocus
    
End Sub

Private Sub Form_Load()
    CentraForm Me
    Screen.MousePointer = 1
    Set tbl_tipodoc = bd_sgp.OpenTable("tipo_documentos")
    
End Sub


Private Sub LimpaTela()
    txt_nome = ""

End Sub




Private Sub lbo_tipodoc_Click()
    AtualTela
    If txt_nome <> "" Then
        opcao = 3
        tbl_tipodoc.Index = "id_codigo"
        tbl_tipodoc.Seek "=", lbo_tipodoc.ItemData(lbo_tipodoc.ListIndex)
        l_codigo = lbo_tipodoc.ItemData(lbo_tipodoc.ListIndex)
        g_tipodoc = l_codigo
    End If
    
End Sub

Private Sub lbo_tipodoc_DblClick()
    opcao = 3
    txt_nome.SetFocus
    l_codigo = lbo_tipodoc.ItemData(lbo_tipodoc.ListIndex)
    g_tipodoc = l_codigo
    
End Sub

Private Sub lbo_tipodoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        AtualTela
        If txt_nome <> "" Then
            If MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro.") = 6 Then
                tbl_tipodoc.Index = "id_codigo"
                tbl_tipodoc.Seek "=", lbo_tipodoc.ItemData(lbo_tipodoc.ListIndex)
                tbl_tipodoc.Edit
                tbl_tipodoc.Delete
                LimpaTela
                PreencheLboBanco
            End If
        End If
    End If
    If KeyCode = 45 Then
        LimpaTela
        opcao = 1
        ProximoCodigo
        txt_nome.SetFocus
    End If

End Sub



Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If

End Sub


