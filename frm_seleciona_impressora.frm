VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_seleciona_impressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione a Impressora a ser utilizada"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   -600
      TabIndex        =   0
      Top             =   -15
      Width           =   6720
      _Version        =   65536
      _ExtentX        =   11853
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   4
      BevelOuter      =   1
      Begin VB.ComboBox cbl_impressora 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frm_seleciona_impressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Definir a impressora padrao
'                frm_seleciona_impressora.Show 1
               'Set Printer = Printers(gImpressoraDefault)



Private Sub cbl_impressora_GotFocus()
    cbl_impressora.Text = cbl_impressora.List(0)
    SendMessageLong cbl_impressora.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub

Private Sub cbl_impressora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gIDImpressoraPadrao = cbl_impressora.ItemData(cbl_impressora.ListIndex)
        gTipoImpressoraSelecionada = "NORMAL"
        gNomeImpressoraSelecionada = cbl_impressora.Text
        Set Printer = Printers(gIDImpressoraPadrao)
        Unload Me
    End If
End Sub
'Private Sub chk_grafica_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cbl_impressora.SetFocus
'End If
'End Sub

Private Sub Form_Activate()
'chk_grafica.SetFocus
End Sub

Private Sub Form_Load()
    PreencheCboImpressora
End Sub
Public Sub MostraMensagens(ByVal pTitulo As String, ByVal pMensagem As String, ByVal pSuperior As Currency, ByVal pEsquerda As Currency, ByVal pLargura As Currency, ByVal pAltura As Currency)
    Me.Top = pSuperior
    Me.Left = pEsquerda
    Me.Width = pLargura
    Me.Height = pAltura
    
    ''CentraForm Me
    'lblTitulo.Width = pLargura
    'lblTitulo.Top = pAltura * 0.25
    'lblMensagem.Width = pLargura
    'lblMensagem.Top = pAltura * 0.6
    'DoEvents
    'lblTitulo.Caption = pTitulo
    'lblMensagem.Caption = pMensagem
End Sub
Public Sub Finaliza()
    Unload Me
End Sub
Private Sub PreencheCboImpressora()
Dim nCount As Integer
Dim Indice As Integer
Dim Impressora As Printer

nCount = 0

With cbl_impressora
  .Clear
    
    For Each Impressora In Printers
        If Not UCase(Impressora.DeviceName) Like "*TM-T20*" And Not UCase(Impressora.DeviceName) Like "*TM-T8*" And Not UCase(Impressora.DeviceName) Like "*MP-4*" And Not UCase(Impressora.DeviceName) Like "*MP-2*" And Not UCase(Impressora.DeviceName) Like "*MP-1*" Then
            .AddItem Impressora.DeviceName
            .ItemData(.NewIndex) = nCount
            .ListIndex = .NewIndex
        End If
        nCount = nCount + 1
    Next
End With

    
    
End Sub

