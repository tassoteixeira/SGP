VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "msoutl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cadastro_menu 
   Caption         =   "Cadastro de Menu Personalizado"
   ClientHeight    =   6435
   ClientLeft      =   1320
   ClientTop       =   1965
   ClientWidth     =   9195
   Icon            =   "cad_menu.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_menu.frx":030A
   ScaleHeight     =   6435
   ScaleWidth      =   9195
   Begin VB.Frame frmUsuario 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.TextBox txt_codigo_usuario 
         Height          =   285
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin MSAdodcLib.Adodc adodc_usuario 
         Height          =   330
         Left            =   3660
         Top             =   240
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_usuario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcbo_usuario 
         Bindings        =   "cad_menu.frx":0750
         Height          =   315
         Left            =   2220
         TabIndex        =   3
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_usuario"
      End
      Begin VB.Label Label5 
         Caption         =   "&Usuário"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frm_dados 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   750
      Width           =   8955
      Begin MSOutl.Outline outline_menu 
         Height          =   5175
         Left            =   60
         TabIndex        =   6
         Top             =   420
         Width           =   4395
         _Version        =   65536
         _ExtentX        =   7752
         _ExtentY        =   9128
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin MSOutl.Outline outline_programa 
         Height          =   5175
         Left            =   4500
         TabIndex        =   8
         Top             =   420
         Width           =   4395
         _Version        =   65536
         _ExtentX        =   7752
         _ExtentY        =   9128
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "&Programas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   150
         Width           =   4275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "&Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   4275
      End
   End
End
Attribute VB_Name = "cadastro_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lUsuario As Integer
Dim lTipo As String
Dim lMenu As String
Dim lSQl As String
Private Menu As New cMenu
Private Programa As New cPrograma
Private Usuario As New cUsuario
Private RstMenu As adodb.Recordset
Private RstPrograma As adodb.Recordset
Private Sub AtualizaOutlineMenu()
    Dim i As Integer
    Dim i2 As Integer
    outline_menu.Clear
    i = -1
    i = AtualizaOutlineMenu2(i, "Cadastros", "CA")
    outline_menu.Expand(0) = True
    i2 = i
    i = AtualizaOutlineMenu2(i, "Consultas", "CO")
    outline_menu.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlineMenu2(i, "Gráficos", "GR")
    outline_menu.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlineMenu2(i, "Movimentação", "MO")
    outline_menu.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlineMenu2(i, "Relatórios", "RE")
    outline_menu.Expand(i2 + 1) = True
End Sub
Function AtualizaOutlineMenu2(x_i As Integer, x_nome As String, x_tipo As String) As Integer
    x_i = x_i + 1
    outline_menu.AddItem x_nome
    outline_menu.Indent(x_i) = 1
    
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "   SELECT Menu"
    lSQl = lSQl & "     FROM Menu"
    lSQl = lSQl & "    WHERE Usuario = " & lUsuario
    lSQl = lSQl & "      AND Tipo = " & Chr(39) & x_tipo & Chr(39)
    lSQl = lSQl & " ORDER BY Menu"
    
    'Abre RecordSet
    Set RstMenu = New adodb.Recordset
    Set RstMenu = Conectar.RsConexao(lSQl)
    
    Do Until RstMenu.EOF
        outline_menu.AddItem Trim(RstMenu("Menu").Value)
        x_i = x_i + 1
        outline_menu.Indent(x_i) = 2
        RstMenu.MoveNext
    Loop
    RstMenu.Close
    Set RstMenu = Nothing
    AtualizaOutlineMenu2 = x_i
End Function
Private Sub AtualizaOutlinePrograma()
    Dim i As Integer
    Dim i2 As Integer
    outline_programa.Clear
    i = -1
    i = AtualizaOutlinePrograma2(i, "Cadastros", "CA")
    outline_programa.Expand(0) = True
    i2 = i
    i = AtualizaOutlinePrograma2(i, "Consultas", "CO")
    outline_programa.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlinePrograma2(i, "Gráficos", "GR")
    outline_programa.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlinePrograma2(i, "Movimentação", "MO")
    outline_programa.Expand(i2 + 1) = True
    i2 = i
    i = AtualizaOutlinePrograma2(i, "Relatórios", "RE")
    outline_programa.Expand(i2 + 1) = True
End Sub
Function AtualizaOutlinePrograma2(x_i As Integer, x_nome As String, x_tipo As String) As Integer
    x_i = x_i + 1
    outline_programa.AddItem x_nome
    outline_programa.Indent(x_i) = 1
    
    'Prepara SQL
    lSQl = ""
    'lSQl = lSQl & "SELECT Tipo, [Nome Interno], [Nome para Menu], [Nome no Disco]"
    'lSQl = lSQl & ", Configuravel, Observacao, Codigo"
    lSQl = lSQl & "   SELECT [Nome para Menu]"
    lSQl = lSQl & "     FROM Programa"
    lSQl = lSQl & "    WHERE Tipo = " & Chr(39) & x_tipo & Chr(39)
    lSQl = lSQl & " ORDER BY [Nome para Menu]"
    
    'Abre RecordSet
    Set RstPrograma = New adodb.Recordset
    Set RstPrograma = Conectar.RsConexao(lSQl)
    
    Do Until RstPrograma.EOF
        outline_programa.AddItem Trim(RstPrograma("Nome para Menu").Value)
        x_i = x_i + 1
        outline_programa.Indent(x_i) = 2
        RstPrograma.MoveNext
    Loop
    RstPrograma.Close
    Set RstPrograma = Nothing
    AtualizaOutlinePrograma2 = x_i
End Function
Private Sub ExcluiMenu(x_nome_tipo As String, x_nome_menu As String)
    If Menu.ExisteProgramaUsuario(lUsuario, x_nome_tipo, x_nome_menu) Then
        If (MsgBox("A opção " & Chr(34) & x_nome_menu & Chr(34) & "," & Chr(10) & "será excluída do menu." & Chr(10) & Chr(10) & "Confirma?", 4 + 32 + 256, "Inclusão de Registro!")) = 6 Then
            If Menu.Excluir(lUsuario, x_nome_tipo, x_nome_menu) Then
                AtualizaOutlineMenu
            Else
                MsgBox "Registro não foi excluído!", vbInformation, "Erro de Integridade"
            End If
        End If
    Else
        MsgBox "Registro inexistente.", vbInformation, "Erro de Integridade!"
        outline_programa.SetFocus
    End If
End Sub
Private Sub Finaliza()
    Set Menu = Nothing
    Set Programa = Nothing
    Set Usuario = Nothing
    frm_cadastro.Show
End Sub
Private Sub IncluiMenu(x_nome_tipo As String, x_nome_menu As String)
    If Programa.LocalizarTipoNome(x_nome_tipo, x_nome_menu) Then
        If Menu.ExisteProgramaUsuario(lUsuario, x_nome_tipo, x_nome_menu) Then
            MsgBox "Programa já está relacionado no menu.", vbInformation, "Registro já Existente!"
            outline_programa.SetFocus
        Else
            If (MsgBox("O programa " & Chr(34) & x_nome_menu & Chr(34) & "," & Chr(10) & "será incluído no menu." & Chr(10) & Chr(10) & "Confirma?", 4 + 32 + 256, "Inclusão de Registro!")) = 6 Then
                Menu.Usuario = lUsuario
                Menu.Tipo = Programa.Tipo
                Menu.Menu = Programa.NomeparaMenu
                Menu.Disco = Programa.NomenoDisco
                Menu.Interno = Programa.NomeInterno
                If Menu.Incluir Then
                    AtualizaOutlineMenu
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
                End If
            End If
        End If
    Else
        MsgBox "Programa não cadastrado.", vbInformation, "Erro de Integridade!"
        outline_programa.SetFocus
    End If
End Sub
Private Sub MarcaOutlineMenu()
    Dim i As Integer
    If outline_menu.ListIndex > -1 Then
        If outline_menu.Indent(outline_menu.ListIndex) = 2 Then
            lMenu = outline_menu.Text
            lTipo = outline_menu.FullPath(outline_menu.ListIndex)
            For i = 1 To 100
                If i = Len(lTipo) Then
                    Exit For
                End If
                If Mid(lTipo, i, 1) = "\" Then
                    lTipo = Mid(lTipo, 1, i - 1)
                    If lTipo = "Cadastros" Then
                        lTipo = "CA"
                    ElseIf lTipo = "Consultas" Then
                        lTipo = "CO"
                    ElseIf lTipo = "Gráficos" Then
                        lTipo = "GR"
                    ElseIf lTipo = "Movimentação" Then
                        lTipo = "MO"
                    ElseIf lTipo = "Relatórios" Then
                        lTipo = "RE"
                    End If
                    Call ExcluiMenu(lTipo, lMenu)
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub MarcaOutlinePrograma()
    Dim i As Integer
    If outline_programa.ListIndex > -1 Then
        If outline_programa.Indent(outline_programa.ListIndex) = 2 Then
            lMenu = outline_programa.Text
            lTipo = outline_programa.FullPath(outline_programa.ListIndex)
            For i = 1 To 100
                If i = Len(lTipo) Then
                    Exit For
                End If
                If Mid(lTipo, i, 1) = "\" Then
                    lTipo = Mid(lTipo, 1, i - 1)
                    If lTipo = "Cadastros" Then
                        lTipo = "CA"
                    ElseIf lTipo = "Consultas" Then
                        lTipo = "CO"
                    ElseIf lTipo = "Gráficos" Then
                        lTipo = "GR"
                    ElseIf lTipo = "Movimentação" Then
                        lTipo = "MO"
                    ElseIf lTipo = "Relatórios" Then
                        lTipo = "RE"
                    End If
                    Call IncluiMenu(lTipo, lMenu)
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub PosicionaUsuario()
    Dim i As Integer
    lUsuario = g_usuario
    txt_codigo_usuario = lUsuario
    If Usuario.LocalizarCodigo(lUsuario) Then
        dtcbo_usuario.BoundText = lUsuario
    Else
        MsgBox "Usuário não cadastrado.", vbInformation, "Atenção!"
        dtcbo_usuario.BoundText = ""
    End If
End Sub
Private Sub dtcbo_usuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        outline_menu.SetFocus
    End If
End Sub
Private Sub dtcbo_usuario_LostFocus()
    If dtcbo_usuario.BoundText <> "" Then
        txt_codigo_usuario.Text = Format(Val(dtcbo_usuario.BoundText), "##")
        lUsuario = Val(dtcbo_usuario.BoundText)
        outline_menu.SetFocus
    Else
        txt_codigo_usuario.Text = ""
    End If
    AtualizaOutlineMenu
End Sub
Private Sub Form_Activate()
    PosicionaUsuario
    AtualizaOutlineMenu
    AtualizaOutlinePrograma
    outline_programa.SetFocus
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    lUsuario = g_usuario
    
    adodc_usuario.ConnectionString = gConnectionString
    adodc_usuario.RecordSource = "SELECT Codigo, Nome FROM Usuario WHERE Situacao = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome"
    adodc_usuario.Refresh
End Sub
Private Sub outline_menu_DblClick()
    MarcaOutlineMenu
End Sub
Private Sub outline_menu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaOutlineMenu
    End If
End Sub
Private Sub outline_programa_DblClick()
    MarcaOutlinePrograma
End Sub
Private Sub outline_programa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaOutlinePrograma
    End If
End Sub
Private Sub txt_codigo_usuario_GotFocus()
    txt_codigo_usuario.SelStart = 0
    txt_codigo_usuario.SelLength = Len(txt_codigo_usuario.Text)
End Sub
Private Sub txt_codigo_usuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        outline_programa.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_usuario_LostFocus()
    Dim i As Integer
    If Val(txt_codigo_usuario.Text) > 0 Then
        If Usuario.LocalizarCodigo(Val(txt_codigo_usuario.Text)) Then
            If Usuario.Situacao = "I" Then
                MsgBox "O usuário " & Usuario.Nome & ", está inativo.", vbInformation, "Usuário não aceito!"
            Else
                dtcbo_usuario.BoundText = Usuario.Codigo
                lUsuario = Usuario.Codigo
                dtcbo_usuario_LostFocus
                Exit Sub
            End If
        Else
            MsgBox "Usuário não cadastrado.", vbInformation, "Atenção!"
        End If
        dtcbo_usuario.BoundText = ""
        txt_codigo_usuario.SetFocus
    Else
        dtcbo_usuario.SetFocus
    End If
End Sub
